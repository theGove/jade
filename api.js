/**
 * An Object for handling Cross origin i-frame communication
 */
const Slick = {
    /**
     * Adds and listener for the message event. You should only call this method once.
     * @param {(event: MessageEvent<*>)=> Promise<Object.<string, any>>} handlerCallback 
     */
    createReceiver(handlerCallback) {
      window.addEventListener("message",
        async event => {
          try {
            if (event.source) {
              log("createReceiver", event)
              //debugger
              const data = await handlerCallback(event)
              event.source.postMessage({ ...data, __slick_id__: event.data.__slick_id__ }, event.origin)
            }// otherwise, there is no source, Gove thinks there is no reason to try to postMessage
          } catch (error) {
            event.source.postMessage({ __slick_error__: error, id: event.data.id }, event.origin)
          }
        }
      );
    },
  
    /**
     * An object that has the post method, used for making cross origin requests across i-frames. 
     */
    requester: function () {
      const requests = {}
      window.addEventListener(
        "message",
        (e) => {
          if (requests[e.data.__slick_id__]) {
            if (e.data.__slick_error__) {
              requests[e.data.__slick_id__].reject(e.data.__slick_error__)
            } else {
              requests[e.data.__slick_id__].resolve(e.data)
            }
          }
        },
      );
      return {
        async post(contentWindow, data, origin = "*") {
          const __slick_id__ = Math.random()
          const promise = new Promise((resolve, reject) => {
            requests[__slick_id__] = { resolve, reject }
          })
          contentWindow.postMessage({ ...data, __slick_id__ }, origin)
          const result = await promise
          delete requests[__slick_id__]
          return result
        }
      }
    }()
  }
  
  async function initialize_api() {
    // we are opening the api
    Slick.createReceiver(async event => {
      if (event.data.mode === "get-page") {
        // this is a request for a page
        log("at init in blogger")
        let source = null
        if (event.data.webPath) {
          source = await get_page_content({ webPath: event.data.webPath })
        } else if (event.data.url) {
          source = await get_page_content({ url: event.data.url })
        } else {
          return "get-page requested but neither webPath nor url provided"
        }
        return { source }
      }
      return {
        message: "Blogger API: External event. No need to process result",
      }
    })
    document.body.replaceChildren(window.location.host)
  }
  
  
  // return the id of a blogger post based on its page_path
  async function bloggerId(page_path) {
    const msgUint8 = new TextEncoder().encode(page_path); // encode as (utf-8) Uint8Array
    const hashBuffer = await crypto.subtle.digest("SHA-1", msgUint8); // hash the page_path
    const hashArray = Array.from(new Uint8Array(hashBuffer)); // convert buffer to byte array
    const hashHex = hashArray
      .map((b) => b.toString(16).padStart(2, "0"))
      .join(""); // convert bytes to hex string
    let hash36 = ""
    // encode blocks of 10 characters at a time to base 36
    for (let x = 0; x < 4; x++) {
      hash36 += parseInt(hashHex.substring(x * 10, x * 10 + 10), 16).toString(36)
    }
    return hash36;
  }
  
  
  
  
  
  
  async function get_page_content(parameters) {
    log("parameters-->", parameters)
    const params = parameters
    const page_url = new URL(window.location)
    let response = null
    let api_fetch_url = null
    if (params.url) {
      api_fetch_url = params.url
    } else {
      api_fetch_url = `${page_url.protocol}//${page_url.host}/2022/02/${params.webPath}.html`
    }
  
    response = await fetch(api_fetch_url)
    const page_content = await response.text()
    return page_content
  }
  
  function log(...args) {
    return
    console.log(window.location.href, ...args)
  }
  
  async function api_request(message_object, api_url) {
    //call the specified api, creating the iframe if it is the first call to this api
    //api url is the domain name where where the api is located.
    //debugger
    //console.log("api_request", message_object, api_url)
    const frame_id = message_object.source
    let iframe = document.getElementById(frame_id)
  
    // if the iframe has successfully received a response from the API, just use it
    if (iframe?.dataset?.status === "verified") {
      return await Slick.requester.post(iframe.contentWindow, message_object, "*")
    }
  
    // if we just created the iframe, the handler from it's content will not have loaded, set up
    // a race to see if the response comes back with in 100 milliseconds, if not try again
  
    iframe = document.createElement("iframe")
    iframe.id = frame_id
    iframe.src = api_url
    iframe.style.display = "none"
    document.body.append(iframe)
  
  
    let result = null
    let delay = 100
    while (!result) {
      delay = delay * 1.5  // increase the time we wait  by 50% for each iteration in case we are on a slow connection
      //log ("waiting",delay,"milliseconds" )
      result = await Promise.race([
        Slick.requester.post(iframe.contentWindow, message_object, "*"),
        new Promise((resolve, reject) => {
          let wait = setTimeout(() => {
            clearTimeout(wait);
            resolve(null);
          }, delay)
        })
      ])
    }
    iframe.dataset.status = "verified"
    return result
  }
  
  // end of api functions.  functions below are useful when this is the only module loaded in blogger

  function jade_decode(text){
    // does a base64 reverse decode
    const iqTag="<"+"!--JADE:data-->"
    return atob(text.split(iqTag)[1].split("").reverse().join(""))
  }

  async function getBookData(initialToc) {
    const data = jsyaml.load(initialToc)
    //console.log(data)
    if (data.tocUrl) {
      let source = null
      if (data.tocUrl.includes(".blogspot.com")) {
        // fetch through blogger and an api iframe
        const prefix = data.tocUrl.split(".blogspot.com")[0]
        const api_response = await api_request({ mode: "get-page", url: data.tocUrl }, `${prefix}.blogspot.com?api`)
        source = getRealContentFromBlogPost(api_response.source)
      } else {
        // it's not a blogger location, let's hope it allows for a CORS request...
        const response = await fetch(data.tocUrl)
        source = await response.text()
      }
      return jsyaml.load(source)
    }
    return data
  }
  
  
  function getRealContentFromBlogPost(source) {
    let metadata = {}
    let result = null
  
    if (source.includes("<" + "!--JADE:data-->")) {
      result = jade_decode(source)
    } else {
      // not published through normal route, use normal tags from blogger template to parse.  This is dangerous because blogger could change them
      result = source.split("<" + 'div class="post-body"><div style="clear:both;"></div>')[1].split("<" + 'div style="clear:both; padding-bottom')[0]
    }
  
    return result
  
  }
  
  