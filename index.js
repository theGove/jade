"use strict"
let PLATFORM=null//used to be able to know if we are in excel or word or powerpoint throughout the code
const jade_settings={}
let jade_css_suffix=""
const jade_panels=['panel_home','panel_examples']
const jade_panel_labels=["Home", "Examples", "Output"]
const jade_code_panels=[]
const jade_panel_stack=['panel_home']
const jade_imports={} // used to keep track of which modules have been imported so we don't try to import twice
const jade_public={}
const jade_modules = {
  add(mod,...arr){
    //console.log("at jade_modules.add")
    if(!this[mod]){this[mod]={}}
    //console.log("arr",arr)
    //console.log("mod",mod)
      for(const func of arr){
         //console.log("func",func)
         this[mod][func.name] = func
      }
  }
}


class Jade{

  // Class Properties

  // Class Methods exposed to be called by Jade Users
  static automate(fn){
    return (...event) => Excel.run((excel)=>fn(excel,...event))
  }

  static async load_gist(gist_id, module_name){
  // We can use gist as repository for the code then either
  // consume it or import it.  Gists have multiple files.
  // .js files become modules in jsvba
  // other file types are stored in window.gist_files object
  // using their filename as the key

  // if module_name is null, load at global scope
  // if module_name is specified, use that name as the module name
  // otherwise, use the filename from the gist as the module name
  
  
  //console.log("gisting", gist_id)

    //Check to see if we are running on local host.  If so, we'll check to see if we can get the gist locally
    if(window.location.hostname==="localhost"){
      // this is implementing the same logic as load_gist_from_local_server.  Gove's not sure where
      // load_gist_from_local_server is called.  This is desigend to work the VITE local server configured
      // for local development
      const url=window.location.origin + "/local-gist-server"
      const response = await fetch(`${url}/${gist_id}/manifest.json`)
      if(response.status===200){
        let files = await response.json()
        const data={files:{}}
        for(const file of files){
          const response = await fetch(`${url}/${gist_id}/${file}`)
          data.files[file]={
            filename:file,
            content:await response.text()
          }
        }
        jade.integrate_gist(gist_id,data, module_name)  
        return
      }
    }


    try{
      
     //console.log("gist_id",gist_id)  
      
      const gist_limit_personal_access_token = localStorage.getItem('gist_limit_personal_access_token')

      //first, check to see if we have a cached version of the gist   
      let data = await jade.read_object_from_workbook("gist:" + gist_id)
     //console.log("gist_cache",data)
      
      let cached_gist_current=false // a variable to tell us if we need to fetch the gist

      if(Object.keys(data).length > 0){ 
        // there is a cached version of this gist, let's check to see if it is current
       //console.log("----> we have a cached version of gist:", data.etag)
        const headers={'If-None-Match': data.etag }

        if(gist_limit_personal_access_token){
            headers.Authorization = "token " + gist_limit_personal_access_token
        }

        const options = {
          method: "HEAD",
          headers: headers
        };
        let response = await fetch(`https://api.github.com/gists/${gist_id}?${Date.now()}`,options)  

        if(response.status===304){
          // the cached gist is identical to the one on the server.  We already have it.
          cached_gist_current = true
         //console.log("x-ratelimit-remaining", response.headers.get("x-ratelimit-remaining"))
         //console.log("----> using cached version of gist")
        }
      }else{
       //console.log("----> no cached version of gist")
      }

      let response
      if(!cached_gist_current){
       //console.log("----> fetching new version of gist")
        // need to get a fresh copy of the gist, either we have no copy, or our copy is outdated
        const options = {method: "GET"}

        // if a gist_limit_personal_access_token has been specified in the local storage, use it for 
        // purposes of increasing the rate limit
        if(gist_limit_personal_access_token){
          options.headers = {
            'Authorization': "token " + gist_limit_personal_access_token
          }
        }

        response = await fetch(`https://api.github.com/gists/${gist_id}`, options)

        //console.log("load gist response", response)
        //console.log("etag", response.headers.get("etag"))
        //console.log("last_modified", response.headers.get("last-modified"))
        //console.log("x-ratelimit-limit", response.headers.get("x-ratelimit-limit"))
       //console.log("x-ratelimit-remaining", response.headers.get("x-ratelimit-remaining"))
        if(response.headers.get("x-ratelimit-remaining")===0){
          // we are out of request capacity.
         //console.log("We are out of gist capacity")
          // need to fall back to atals-query gist server on git hub
          // to publish a gist to this gist-server, run publish <gist-id>
          // from jsvba/gist-server  repo
          response = await fetch(`https://jade-addin.github.io/public-gist-server/${gist_id}.json`)
        }

        ////console.log("x-ratelimit-reset", response.headers.get("x-ratelimit-reset"))
        data = await response.json()

        if(data.files){
          // store the gist file so a subsequent request will not need to ask for it again
          // becuase we make an initial call using if-none-match header, we can know if the
          // content has changed without counting agains the request limits, so we can always
          // the most recent version of the gist.

          const etag=response.headers.get("etag").split("/")[1]
          jade.save_object_to_workbook({// no need to await because we are not waiting on it
            etag:etag,
            files:data.files
            }, 
            "gist:" + gist_id
          )  
        }else{
          console.error("Gist contains no files",data)
          ;console.log("etag", response.headers.get("etag"))
          ;console.log("last_modified", response.headers.get("last-modified"))
          ;console.log("x-ratelimit-limit", response.headers.get("x-ratelimit-limit"))
          ;console.log("x-ratelimit-remaining", response.headers.get("x-ratelimit-remaining"))
          ;console.log("x-ratelimit-reset", response.headers.get("x-ratelimit-reset"))
          return
        }

      }


      jade.integrate_gist(gist_id,data, module_name)  
      
    }catch(e){
     ;console.error("Error fetching gist", e)
     throw "Unable to process Gist"
    }
  }

  static async load_gist_from_local_server(gist_id, module_name, url){
    // behaves like load gist, but it pulls the data from the local "go live" server
    // from vs code.  Open vs code from the folder containing the folder with the gist id
    // the gist must have been cloned locally  

    
    try{
      
     //console.log("=============================================================")  
     //console.log("at load_gist_from_local_server, gist_id:",gist_id)  
     //console.log("URL:",`${url}/${gist_id}/manifest.json`)  
     //console.log("=============================================================")  

     //let response = await fetch(`${url}/${gist_id}`)
     let response = await fetch(`${url}/${gist_id}/manifest.json`)
     let files = await response.json()
      const data={files:{}}
      for(const file of files){
        const response = await fetch(`${url}/${gist_id}/${file}`)
        data.files[file]={
          filename:file,
          content:await response.text()
        }
      }
      jade.integrate_gist(gist_id,data, module_name)  
    }catch(e){
      ;console.error("Error fetching gist", e)
      throw "Unable to process Gist"
    }
  }
  
  static async integrate_gist(gist_id, data, module_name){
    try{// now process the gist

      //global variable
      if(!window.gist_files){
        window.gist_files={}
      }
      window.gist_files[gist_id]={}

      //load non js file first
      for(const file of Object.values(data.files)){
        if(file.filename.slice(-3)!==".js"){
          gist_files[gist_id][file.filename] = file.content
        }
      }

      //process the JS files
      const fn_gist_id=`
      function my_gist_id(){return "${gist_id}"}
      function gist_files(file_name, gist_id="${gist_id}"){
        return window.gist_files[gist_id][file_name]
      }
      ` // gets prepended to each js file so it knows its own gist ID
      for(const file of Object.values(data.files)){
        if(file.filename.slice(-3)===".js"){
          // check syntax of js file
          const parsed_code=Jade.parse_code(file.content)

          if(parsed_code.error){
            alert(parsed_code.error, "Syntax Error in " + file.filename)
            return 
          }

          if(module_name===null){
            await Jade.incorporate_code(file.content + fn_gist_id)
            try{
              auto_exec()
              auto_exec=null// in case a later module gets loaded in the global scope, don't run an old auto_exec
            }catch(e){
              if(e.name!=="ReferenceError"){
              console.error("Error running auto_exec()",e.name)
              }
            }
            //Jade.incorporate_code("auto_exec=null")
          }else if(!module_name){
            const module_name=Jade.file_name_to_module_name(file.filename)
            await Jade.incorporate_code(file.content + fn_gist_id, module_name)
            try{
              jade_modules[module_name].auto_exec()
            }catch(e){
              console.error("Error running auto_exec()",e)
            }
          }else{
            //console.log("module_name", module_name)
            //console.log("file.content", file.content)
            await Jade.incorporate_code(file.content + fn_gist_id, module_name)
            try{
              jade_modules[module_name].auto_exec()
            }catch(e){
              console.error("Error running auto_exec()",e)
            }
          }
        }
      }
    }catch(e){
      ;console.error("Error loading gist",e)
      throw "Unable to process Gist"
    }

  }

  static async load_js(url, module_name){
    // if module_name is null, load at global scope
    // if module_name is specified, use that name as the module name
    // otherwise, use the filename from the gist as the module name
    
  
    //console.log("gisting", gist_id)
      try{
          
        const response = await fetch(`${url}?${Date.now()}`)
        const code = await response.text()
        try{
          if(module_name===null){
            await Jade.incorporate_code(code)
            try{
              auto_exec()
              auto_exec=null// in case a later module gets loaded in the global scope, don't run an old auto_exec
            }catch(e){
              if(e.name!=="ReferenceError"){
              console.error("Error running auto_exec()",e.name)
              }
            }
            //Jade.incorporate_code("auto_exec=null")
          }else if(!module_name){
            const module_name=Jade.file_name_to_module_name(url.substring(url.lastIndexOf('/')+1))
            await Jade.incorporate_code(code, module_name)
            try{
              jade_modules[module_name].auto_exec()
            }catch(e){
              console.error("Error running auto_exec()",e)
            }
          }else{
            //console.log("module_name", module_name)
            //console.log("file.content", file.content)
            await Jade.incorporate_code(code, module_name)
            try{
              jade_modules[module_name].auto_exec()
            }catch(e){
              console.error("Error running auto_exec()",e)
            }
          }
          
        }catch(e){
          ;console.error("Error loading JS",e)
        }
        
      }catch(e){
       ;console.error("Error fetching JS", e)
      }
    }
  
    static async getBloggerPost(blogName, postName, pubYear, pubMonth){

      // get a blog post.  if pubYear and pubMonth not supplied, Assumes published on 2022/02
      const blogPathElements = blogName.split("/")
      if(blogPathElements.length===4){
        // we have a full path specified as the first argument, overwrite the others
        blogName = blogPathElements[0]
        pubYear  = blogPathElements[1]
        pubMonth = blogPathElements[2]
        postName = blogPathElements[3]
      }
      
      //console.log("blogPathElements",blogPathElements)
      if(window.location.hostname==="localhost"){
        // running on localhost, get from localserver
        
        const url=window.location.origin + "/local-blog-server"
        let response = await fetch(`${url}/${blogName}/metadata/${postName}.json`)
        if(response.status===200){
          //found metadata for the post
          const metadata = await response.json()
          response = await fetch(`${url}/${blogName}/posts/${postName}.${metadata.type}`)
          if(response.status===200){
            // ;console.log("------ Got post from local blog server ------\n",blogName,postName,"\n--------------------------------------------")
            return await await response.text()
          }  
        }
      }



      //if we got this far, we need to fetch from blogger
      // get from blogger
      const request = { mode: "get-page" }
      request.source = blogName
      request.webPath = postName
      request.pubYear = pubYear
      request.pubMonth = pubMonth
      //console.log("request",request)
      let api_url = `https://${request.source}.blogspot.com?api`
      let api_response = await api_request(request, api_url)
      let source = getRealContentFromBlogPost(api_response.source)
      return source

    }

    static async isHex(rawData){
      // returns true if each character in rawData is in the range 0-F
      for(const digit of rawData.toUpperCase().split("")){
        if(isNaN(digit) && (digit.charCodeAt(0)<65 || digit.charCodeAt(0)>70)){
          return false
        }
      }
      return true
    }

    static async bindModule(url_or_code){
      // to invoke this code, import a code module with this code: excel-assessment/post/100
      // where excel-assessment is the name of the blog (before .blogspot.com) and 
      // post is the name of the post.  The blog must be published in february of 2022 so it has a url of 
      // https://excel-assessment.blogspot.com/2022/02/post.html
      // anything else in the path will be sent as parameters to a function named "autoexec"

      // special loading to handle grading of excel assessments
      //console.log("I'm Binding!")

      if(url_or_code.length===0){
        if(window.location.hostname==="localhost"){
          tag("module-url-or-code").value="excel-assessment|2024/02/is-110-fall-2024|wscxyVdVDRP8oK3xhT3dlbnxMdW50"
        }
        return
      }  
      
      tag('spinner').style.display=''

      let source=null

      if(tag("module-url-or-code").value){
        localStorage.setItem("JadeImportCode",tag("module-url-or-code").value)
      }

      //console.log("is hex",await Jade.isHex(url_or_code))
      if(url_or_code.substring(0,7)==="http://" || url_or_code.substring(0,8)==="https://"){
        // this is a url to a js file
        const response = await fetch(url_or_code);
        source = await response.text();
        // still needs work
        return
  
      }else if(url_or_code.length >= 32 && await Jade.isHex(url_or_code)){
        // this could be a gist_id.  Let's check
        const response = await fetch("https://api.github.com/gists/");
        const movies = await response.text();
        //console.log(movies); 
        // still needs work
        return
      }

      // currently, only binding to blogger module.  need to add url, gist, and internal module (written in jade)

      let parameters=url_or_code.split("|")      
      const parts=parameters.shift().split("/")
      let blogName = parts[0]
      let postName = null
      let pubYear = "2022"
      let pubMonth = "02"
     
      if(parts.length===1){
        // only supplied the blog identifier.
        postName = "index"
      }else if(parts.length===2){
        // only supplied blog identifed and post name
        postName = parts[1]
      }else if(parts.length===3){
        //console.log("three parts of the blogger identifer supplied.  This is invalid")
        // need to give an error message
      }else if(parts.length===4){
        pubYear = parts[1]
        pubMonth = parts[2]
        postName = parts[3]
      }
      //console.log("parts", parts)
      //console.log("incporporating", blogName,pubYear, pubMonth, postName, parameters)
      const theModule = await Jade.incorporateBloggerModule(blogName,pubYear, pubMonth, postName, parameters)
      const params1={
        blogName,postName,pubMonth,pubYear,
        args:parameters,
        module:theModule
      }
      try{
        // for opened modules the auto-run function is called initialize.
        theModule.initialize(params1)
      }catch(e){
        console.warn("Unable to run initialize() of blog module", blogName. postName,e)
      }


      
  }
  
  static async incorporateBloggerModule(blogName,pubYear, pubMonth, postName){
    //loads a code module from blogger  gives it the name blogname_postname
    // returns a reference to the jade module

    const blogPathElements = blogName.split("/")
    if(blogPathElements.length===4){
      // we have a full path specified as the first argument, overwrite the others
      blogName = blogPathElements[0]
      pubYear  = blogPathElements[1]
      pubMonth = blogPathElements[2]
      postName = blogPathElements[3]
    }
    

    const moduleName = `${blogName}/${pubYear}/${pubMonth}/${postName}`
    if(!jade_modules[moduleName]){
        const source = await Jade.getBloggerPost(blogName, postName, pubYear, pubMonth)
        const incorporateBloggerModule = await Jade.incorporate_code(source, moduleName)  
        //if(saveToWorkbook){Jade.save_module_to_workbook(source, "initialize")}
    }  
    return jade_modules[moduleName]
}

  
  static async import_code_module(url_or_gist_id){
     //console.log("at import code mod", url_or_gist_id)
     let module_name=null
     jade_settings.workbook.module_to_import=url_or_gist_id

      Jade.hide_element("import-module")
      Jade.save_settings()
      if(!url_or_gist_id){
          return
      }
      let url=null

      

      if(url_or_gist_id.substr(0,4)==="http"){
          // check to see if GIST url
          if(url_or_gist_id.toLowerCase().includes("gist.github.com")){
              const url_data = url_or_gist_id.split("/")
              url='https://api.github.com/gists/' + url_data[url_data.length-1] + "?" + Date.now()
          }else{
              url=url_or_gist_id
          }

      }else{
          // this looks like a gist id.  we should probably check it
          // sometime.  for now, let's just assume it is
          url = 'https://api.github.com/gists/' + url_or_gist_id + "?" + Date.now()
      }

      // now we have the URL to process
    
      try{
          const response = await fetch(url)
          var data = await response.text()
        }catch(e){
          alert(e.message,"Error Loading Module")  
          ;console.log("Error fetching gist", e)
          return
        }
      //console.log("data",data)  
      const files=[]  
      // let's see if we got back json
      try{
          var gist_json=JSON.parse(data)
          // there was no error parsing the json, so this must be gist manifest
          // load the individual files
          try{
             //console.log("parsing json from gist")
              for(const file of Object.values(gist_json.files)){
                  files.push({name:file.filename.split(".js")[0], code:file.content})
              }  
          }catch(e){
              alert(e.message, "Error Parsing Gist")
              return
          }
      }catch(e){
         //console.log("module is not json")
          // json was not valid, assume we have js
          // check to see if there is a comment that specifies a module name
          //
          let name = null
          if(data.includes("jade.module:")){ 
            try{
               //console.log("found ace lable")
              name = JSON.parse(data.split("jade.module:")[1].split("*/")[0]).name
            }catch(e){
               //console.log("ace label invalid")
              const url_data=url.split("/")
              name=url_data[url_data.length-1]
            }
          }
          if(!name){
            //either there was no comment to specify a name, or there was an error in reading it
            // we won't overwrite a module unless it is named and the name is something other than "Code"
            // so here, we are going to give it a number to make it unique
             //console.log("no name for you")
            let x=1
            while(!!tag(Jade.panel_label_to_panel_name("Code "+ x ))){x++}
            files.push({name:"Code " + x, code:data})
          }
      }   

      // now we should have files looking like this
      // files:[{name:module1,content:"function zeta(){..."}, {name:module2,content:"function beta(){..."}]
      // we need to add or update based on the name.
      for(const file of files){
         //console.log(file.name,Jade.panel_label_to_panel_name(file.name))
          if(!!tag(Jade.panel_label_to_panel_name(file.name+" Module"))){
              // a module with this name already exists,  update
             //console.log("========= ready to update ============", file.name)
              const editor=ace.edit(Jade.panel_label_to_panel_name(file.name) + "_module-content")
              editor.setValue(file.code)
             //console.log(editor.getValue())
          }else{
              // no module with this name exists, append
              Jade.add_code_module(file.name, file.code)
          }
      }
      
  }
  static set_css(user_css, style_tag_id="user_style"){

    if(style_tag_id==="user_style"){
      // only set the jade_css_suffix style_tag_id is not supplied
      jade_css_suffix=user_css
    }

      const user_style_tag = document.getElementById(style_tag_id)
      if(user_style_tag ){
        user_style_tag.remove()
      }
      document.head.insertAdjacentHTML("beforeend", `<style id="${style_tag_id}">`+ user_css + "</style>")
  
  }

  static add_library(url){
      // adds a JS library to the head section of the HTML sheet
      const library = document.createElement('script');
      library.setAttribute('src',url);
     //console.log("library",library)
      document.head.appendChild(library);
  }
  static close_canvas(){
      jade_panel_stack.pop()
      Jade.show_panel(jade_panel_stack.pop())
  }
  static open_editor(){
      Jade.show_panel(jade_code_panels[0])
  }
  static open_output(){
      Jade.show_panel("panel_output")
  }
  static open_automations(show_close_button, heading, list){
      Jade.show_automations(show_close_button, heading, list)
  }

  static open_tools(){
    //gist id for jade tools (jet)
    jade.load_gist("0ad9cb51c9c7cdd79500a8e51ee85a18")
  }


  static reset(){
      Jade.show_panel("panel_home")
  }
  // static show_html(html){
  //     //A simple function that is mapped differntly for examples than for modules
  //     //this is the module mapping
  //     Jade.open_canvas("html", html)
  // }
  static open_canvas(panel_name, html, show_panel_close_button, style_name){
      //if(style_name){
          Jade.set_style(style_name)
      //}

      if(!tag(panel_name)){
          Jade.build_panel(panel_name)
      }

      if(!jade_panels.includes(panel_name)){
          jade_panels.push(panel_name)
      }

      Jade.show_panel(panel_name)

      if(html){
          if(show_panel_close_button || show_panel_close_button===undefined){
              tag(panel_name).innerHTML=Jade.panel_close_button(panel_name) + html
          }else{
              tag(panel_name).innerHTML= html
          }
          
      }
  }
  static print(data, heading){
      //if(!header && )
      if(!tag("panel_output").lastChild.lastChild.firstChild.tagName && !heading){
          //no output here, need a headdng
          heading=""
      }
      if(heading){
          // there is a header, so make a new block
         //console.log("at data")
          const div = document.createElement("div")
          div.className="jade-output"
          const header = document.createElement("div")
          header.className="jade-output-header"  
          const d = new Date()
          let ampm=" am"
          let hours=d.getHours()
          if(hours >11){
              ampm="pm"
              if(hours>12){
                  hours=hours-12
              }
          }
          header.innerHTML = '<span class="jade-output-time">' + hours + ":" + ("0"+d.getMinutes()).slice(-2) + ":" + ("0"+d.getSeconds()).slice(-2) + ampm + "</span> " + heading + '<div class="jade-output-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>'
          const body = document.createElement("div")
          body.className="jade-output-body"  
          body.innerHTML = '<div style="margin:0;font-family: monospace;">' + data.split("\n").join("<br />")  + "<br />"+ "</div>"
          div.appendChild(header)
          div.appendChild(body)
          tag("panel_output").appendChild(div)
      }else{
          // no header provided, append to most recently added
          tag("panel_output").lastChild.lastChild.firstChild.innerHTML += data.split("\n").join("<br />") + "<br />"
      }

    
  }
  static open_examples(){
      const panel_name="panel_examples"
      Jade.set_style()
      if(!tag("panel_examples").dataset.initialized){
        Jade.fill_examples()
        tag("panel_examples").dataset.initialized=true
      }
      Jade.show_panel(panel_name)
  }

  static async use(name, module_name){
    // imports a gist using a name to find the gist ID
    if(!jade_imports[name]){
      // only import once
     //console.log(jade_settings.workbook)
      const url = jade_settings.workbook.gist_name_server + name + "?a=8"
     //console.log("------------->",url)
      const response = await fetch(url)
      const gist_id = await response.text()
     //console.log(gist_id)
      if(module_name===undefined){module_name=null}
      await Jade.load_gist(gist_id, module_name)
      jade_imports[name]=gist_id
    }
  }

  // Class methods that still need work before shown to the public

  static set_theme(theme_name){
    Jade.set_style(theme_name)
  }
  static list_themes(){
      for(const [theme, url] of Object.entries(jade_settings.workbook.styles)){
         //console.log(theme, url)
      }
  }


  // Class Methods NOT meant to be called by Jade Users
  // It's not really a problem if they do, we just don't
  // think they are useful and we don't document them.


  static officeReady(info){// invoked when the office addin infrastructure has loaded
    //console.log ("host, hosttype----->",info.host, Office.HostType.Excel) 

    tag("open-tools").onclick = function(){Jade.open_tools()};
    tag("open-editor").onclick = function(){Jade.open_editor()};
    tag("add-module-row").onclick = function(){Jade.toggle_element('add-module');tag('new-module-name').focus()};
    tag("module-add-button").onclick = function(){Jade.add_code_module(tag('new-module-name').value)};
    tag("show-import-module-row").onclick = Jade.show_import_module;
    tag("module-import-button").onclick = function(){Jade.import_code_module(tag('gist-url').value)};
    tag("open-examples-row").onclick = Jade.open_examples;
    tag("list-automations-row").onclick = Jade.open_automations;
    tag("show-survey-row").onclick = function(){Jade.toggle_element('survey');tag('fb-type').focus();tag('fb-message').scrollIntoView(true);};
    tag("fb-button").onclick = Jade.submit_feedback;
    tag("show-settings-row").onclick = function(){Jade.toggle_element('settings-page');tag('settings-button').scrollIntoView(true);};
    tag("settings-button").onclick = Jade.save_settings;
    PLATFORM = info.host  // make the host knowable elsewhere    

    if(Jade.get_settings_from_workbook() === null){
      console.log("using default values")
      // no jade_settings so let's configuire some defaults
      jade_settings.user={
        ace_options:{
          selectionStyle:"line",
          highlightActiveLine:true,
          highlightSelectedWord:true,
          readOnly:false,
          copyWithEmptySelection:false,
          cursorStyle:"ace",
          mergeUndoDeltas:true,
          behavioursEnabled:true,
          wrapBehavioursEnabled:true,
          enableAutoIndent:true,
          showLineNumbers:true,
          hScrollBarAlwaysVisible:false,
          vScrollBarAlwaysVisible:false,
          highlightGutterLine:true,
          animatedScroll:false,
          showInvisibles:false,
          showPrintMargin:false,
          printMarginColumn:80,
          printMargin:80,
          fadeFoldWidgets:false,
          showFoldWidgets:true,
          displayIndentGuides:true,
          showGutter:true,
          fontSize:"12pt",
          scrollPastEnd:0,
          theme:"ace/theme/tomorrow",
          maxPixelHeight:0,
          useTextareaForIME:true,
          scrollSpeed:2,
          dragDelay:0,
          dragEnabled:true,
          focusTimeout:0,
          tooltipFollowsMouse:true,
          firstLineNumber:1,
          overwrite:false,
          newLineMode:"auto",
          useWorker:false,
          navigateWithinSoftTabs:false,
          wrap:true,
          indentedSoftWrap:true,
          foldStyle:"markbegin",
          mode:"ace/mode/javascript",
          enableMultiselect:true,
          enableBlockSelect:true,
          tabSize:2,
          useSoftTabs:true
        } 
      }
      jade_settings.workbook={
        code_module_ids:[],
        gist_name_server:"https://gns.jsvba.com/",
        examples_gist_id:"f8e5fc20cff0c19a27765e7ce5fe54fe",
        styles:{
          system:null,
          none:"/*No Theme CSS Used*/",
          mvp:"https://cdnjs.cloudflare.com/ajax/libs/mvp.css/1.8.0/mvp.css",
          marx:"https://cdnjs.cloudflare.com/ajax/libs/marx/4.0.0/marx.min.css",
          water:"https://cdn.jsdelivr.net/npm/water.css@2/out/water.css",
          "dark water":"https://cdn.jsdelivr.net/npm/water.css@2/out/dark.css",
          sajura:"https://unpkg.com/sakura.css/css/sakura.css",
          tacit:"https://cdn.jsdelivr.net/gh/yegor256/tacit@gh-pages/tacit-css-1.5.5.min.css",
          pure:"https://unpkg.com/purecss@2.0.6/build/pure-min.css",
          picnic:"https://cdn.jsdelivr.net/npm/picnic",
          wing:"https://unpkg.com/wingcss",
          chota:"https://unpkg.com/chota@latest",
          bootstrap:"https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css",
        }
      }
    }else{// if xl_settings a null object
      console.log("using stored values")
      //console.log("xl_settings",xl_settings.value)
      const stored_jade_settings = Jade.get_settings_from_workbook()
      jade_settings.workbook = stored_jade_settings.workbook
      jade_settings.user = stored_jade_settings.user
    }// if jade_settings null object
    //console.log("before start_me_up, jade_settings", jade_settings)
    Jade.configure_settings()
    Jade.start_me_up()
    if(document.getElementById("menu-msg")){
      document.getElementById("menu-msg").style.display = "block"
    }
  }
  static async start_me_up(){
  //console.log("at start_me_up")


  const urlParams = new URLSearchParams(window.location.search);
  let mode = urlParams.get('mode')
  //console.log("mode", mode)

  jade_settings.workbook.styles.system=tag("head_style").innerText
  jade_panels.push("panel_home")
  
  //load code from one gist if specified.  
 //console.log("about ot load")
  if(mode==="jade"){
    // an intermediate step to allow jade tools to stay on the main jade screen until new manifest is published (2/2024)
    if(tag("open-tools")){
      tag("open-tools").style.display="none"
    }
  }else if(mode==="tools"){
    // loading excel tools
    jade.load_gist("0ad9cb51c9c7cdd79500a8e51ee85a18")
    mode="loading"
  }else if(mode==="module"){
  //binding a module to this workbook

    let boundModuleExists=false 
    for(const code_module_id of jade_settings.workbook.code_module_ids){
      const module_info = Jade.get_module(code_module_id)
      const module_name= module_info.name //doc.getElementsByTagName("name")[0].textContent
      const module_code = atob(module_info.code) //atob(doc.getElementsByTagName("code")[0].textContent)
      Jade.add_code_editor(module_name, module_code,code_module_id, null, JSON.parse(options))    
      
      Jade.incorporate_code(module_code, module_name)
      //Jade.add_code_editor(module_name, module_code,code_module_id, null, JSON.parse(options))        
    }
    //console.log("boundmodule---->", boundModuleExists)

    if(boundModuleExists){
      tag("panel_home").innerHTML=``  
      //console.log("Earth Earth Earth Earth Earth Earth Earth Earth Earth ")

      jade_modules.initialize.initialize()

    }else{

    tag("panel_home").innerHTML=`
    <div style="background-color:darkgreen; text-align:center;padding:1rem" >
      <div style="background-color:white; padding:1rem;">
        <div style="margin-bottom:1rem">Enter Code:</div>
          <input id="module-url-or-code" type="text" style="width:100%"> 
        <div style="text-align:right">
          <button onClick="Jade.bindModule(tag('module-url-or-code').value)">Import</button>
          </div>
          <img id="spinner" src="assets/spinner.gif" style="display:none">
      </div>
    </div>
    `
    const currentCode = localStorage.getItem("JadeImportCode")
    if(currentCode){
      tag("module-url-or-code").value=currentCode
    }  


    }
  }else if(jade_settings.workbook.load_gist_id){
    //console.log("in if")
    if(jade_settings.workbook.load_gist_id.substr(0,4).toLowerCase()==="http"){
      // must be a javascript file
      Jade.load_js(jade_settings.workbook.load_gist_id)
    }else{
      // must be a gistid
      Jade.load_gist(jade_settings.workbook.load_gist_id)
    }
    mode="loading"    
  }

  if(mode!=="loading"){
    tag("panel_home").style.display="block"
  }
  
  if(!tag("new-module-name")){
    // the ret
    return
  }
  if(mode==="jade" || !mode){  // code that only needs to be run for the main Jade panel
    // add event listner to "add code module" input
    tag("new-module-name").addEventListener("keyup", function(event) {
        event.preventDefault();
        if (event.key === 'Enter') {
            tag("module-add-button").click();
        }
    });

    // add event listner to "import module" input
    tag("gist-url").addEventListener("keyup", function(event) {
      event.preventDefault();
      if (event.key === 'Enter') {
          tag("module-import-button").click();
      }
    });


    // fit the editor to the windows on resize
    window.addEventListener('resize', function(event) {
      //console.log("hi")
      for(const panel_name of jade_code_panels){
        tag(panel_name + "_editor-page").style.height = Jade.editor_height()
      }
    }, true);
    Jade.init_examples()
    Jade.init_output()
    
    // ---------------- Initializing Code Editors -----------------------------
    if(jade_settings.workbook.code_module_ids.length>0){// show the button to view code modules
      Jade.show_element("open-editor")
    }
        for(const code_module_id of jade_settings.workbook.code_module_ids){
          const module_info = Jade.get_module(code_module_id)
          const module_name= module_info.name //doc.getElementsByTagName("name")[0].textContent
          const module_code = atob(module_info.code) //atob(doc.getElementsByTagName("code")[0].textContent)
          const options= module_info.options //atob(doc.getElementsByTagName("options")[0].textContent)
          Jade.add_code_editor(module_name, module_code,code_module_id, null, JSON.parse(options))        
        }
  }// end of code just for the JADE main panel
  }
  static configure_settings(){
   //console.log("jade_settings", jade_settings)
    //if(!tag('settings-page').className.includes("hidden")){
      //tag('jade-theme').focus();
      
      //console.log("fontSize", jade_settings.user.ace_options.fontSize)  
      tag("jade-font-size").value = jade_settings.user.ace_options.fontSize.replace("pt","")
      if(jade_settings.user.ace_options.wrap===false){
        tag("jade-word-wrap").value="no-wrap"
      }else if(jade_settings.user.ace_options.indentedSoftWrap){
        tag("jade-word-wrap").value="wrap-indented"
      }else{
        tag("jade-word-wrap").value="wrap"
      }
      //console.log("theme", jade_settings.user.ace_options.theme)
      //console.log("theme", jade_settings.user.ace_options.theme.split("/")[2])
      tag("examples-gist-id").value  = jade_settings.workbook.examples_gist_id
      tag("jade-theme").value  = jade_settings.user.ace_options.theme.split("/")[2]
      tag("jade-line-numbers").checked = jade_settings.user.ace_options.showGutter
      if(jade_settings.workbook.load_gist_id){
        tag("load-gist-id").value = jade_settings.workbook.load_gist_id
      }
    //}
  }
  static async save_settings(){
    if(tag("examples-gist-id").value && jade_settings.workbook.examples_gist_id!==tag("examples-gist-id").value){
      // different examples specified.  Need to rebuild
      Jade.rebuild_examples(tag("examples-gist-id").value)
    }
    jade_settings.workbook.examples_gist_id=tag("examples-gist-id").value
    jade_settings.workbook.load_gist_id=tag("load-gist-id").value
    jade_settings.user.ace_options.theme="ace/theme/" + tag("jade-theme").value
    jade_settings.user.ace_options.fontSize=tag("jade-font-size").value + "pt"
    jade_settings.user.ace_options.showGutter=tag("jade-line-numbers").checked
    switch(tag("jade-word-wrap").value){
      case "wrap":
        jade_settings.user.ace_options.wrap=true
        jade_settings.user.ace_options.indentedSoftWrap=false
        break
      case "wrap-indented":
        jade_settings.user.ace_options.wrap=true
        jade_settings.user.ace_options.indentedSoftWrap=true
        break
      default:  
        jade_settings.user.ace_options.wrap="off"
    }
    //console.log(jade_settings.user.ace_options)
    Jade.apply_editor_options(jade_settings.user.ace_options)
  
    await Jade.write_settings_to_workbook()
  
    Jade.hide_element('settings-page')
  }
  static async save_object_to_workbook(the_object, the_key){
    Jade.setSetting(the_key, the_object)
  }
  static async read_object_from_workbook(the_key){
    return Jade.getSetting(the_key)
  }

  static async write_settings_to_workbook(){
    await Jade.setSetting('jade_settings', jade_settings)
  }

  static get_settings_from_workbook(){
    return Jade.getSetting('jade_settings')
  }

  static apply_editor_options(options){
    for(const panel_name of jade_code_panels){
      //console.log("updating options on ", panel_name)
      const editor = ace.edit(panel_name + "-content");
      editor.setOptions(options)
    }
  }
  static async submit_feedback(){
    // send feedback to the google form
    const message=[]
    if(tag("fb-type").value===""){message.push("<li>You must indicate the <b>type</b> of your feedback.</li>")}
    if(tag("fb-text").value===""){message.push("<li>You must provide the <b>text</b> of your feedback.</li>")}
    if(tag("fb-platform").value===""){message.push("<li>You must provide the <b>platform</b> you are using.</li>")}
    if(message.length===0){
      // ready to submit
      const form_data=[]
      form_data.push("entry.1033992853=")
      form_data.push(encodeURIComponent(tag("fb-type").value))
      form_data.push("&entry.482647522=")
      form_data.push(encodeURIComponent(tag("fb-platform").value))
      form_data.push("&entry.1230850697=")
      form_data.push(encodeURIComponent(tag("fb-email").value))
      form_data.push("&entry.1009690762=")
      form_data.push(encodeURIComponent(tag("fb-text").value))
      form_data.push("&entry.64124153=")
      form_data.push(encodeURIComponent("Addin"))
  
      tag("fb-message").innerHTML=""
      
      const options = {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: form_data.join("")
      };
      //console.log("form data",form_data.join(""))
      try{
      const response = await fetch("https://docs.google.com/forms/u/0/d/e/1FAIpQLSeRCZKHMU0HvXEiia8dn6hL0DfGoAH-PCJZolfiN1S_O90eSQ/formResponse", options)
      //console.log("reson",response.status)
      const data = await response.text()
      //console.log("data",data)
      tag("fb-message").style.color="green"
  
      //console.log("=====================",tag("fb-type").value)
      switch(tag("fb-type").value){
        case "Feature Resuest":
          message.push("Thanks for your feedback. Were not sure when we'll be updaing next but thanks for helping us understand your needs.")
          break
        case "Report Issue":
          message.push("Thanks for reporing this issue. While we cannot respond to every problem report, we'll do what we can.")
          break
        case "Offer to help with development":
          message.push("Thanks for offering to help on this project.  We'll take a look at your comment and get back to you--if you provided a valid email address.")
          break
        case "Praise for Addin":
          message.push("Thanks a million. We thrive on positive feedback!")
          break
        default:  
          message.push("Thanks for asking this question. While we cannot respond to every question, we'll do what we can.")
      }
      tag("fb-message").innerHTML=message.join("")
      tag("fb-message").scrollIntoView(true);
      setTimeout(function() {
        Jade.hide_element('survey')
        tag("fb-message").innerHTML=""
      }, 10000);
      }catch(e){
        //console.log("form error: ",e)
        tag("fb-message").style.color="red"
        tag("fb-message").innerHTML='Oops.  It looks as though there was a network error.  Your can tray again later or submit at our <a href="https://docs.google.com/forms/d/e/1FAIpQLSeRCZKHMU0HvXEiia8dn6hL0DfGoAH-PCJZolfiN1S_O90eSQ/viewform?usp=pp_url&entry.64124153=web" target="_blank">Google Form<a>.'  
        window.scrollTo(0,document.body.scrollHeight);
        tag("fb-message").scrollIntoView(true);
      }
      
    }else{
      tag("fb-message").style.color="red"
      tag("fb-message").innerHTML="<ul>" + message.join("") + "</ul>"
      tag("fb-message").scrollIntoView(true);
    }
    
  }
  static add_code_module(name,code){
    // a module built with whatever code is in Jade.default_code
    
    if(!name){name="code"}
    // check for duplicate name--that wreaks havoc
    let found_panel=false
    for(const panel_name of jade_code_panels){
     //console.log("panel_name",panel_name,Jade.panel_label_to_panel_name(name))
      if(panel_name === Jade.panel_label_to_panel_name(name) + "_module"){
        //we have a match, and that's a no-no
        alert('A module named "'+name+'" already exists in this workbook.  <br><br>Choose a differnt name.',"Invalid Module Name")
        return
      }
    }


    if(!code){// no code is pased in, determine which default code to import
      if(jade_code_panels.length === 0){
        code = Jade.default_code()
      }else{  
        code = Jade.default_code("panel_" + name.toLowerCase().split(" ").join("_") + "_module")
      }
    }



    Jade.add_code_editor(name, code, "")
    Jade.hide_element("add-module")
    Jade.show_element("open-editor")
    Jade.show_panel(jade_code_panels[jade_code_panels.length-1])
    Jade.write_module_to_workbook(code,jade_code_panels[jade_code_panels.length-1])
  }
  static async get_style(style_name, url, integrate_now){

    if(integrate_now===undefined){
       integrate_now=true
    }
  
    if(!jade_settings.workbook.styles[style_name]){
      jade_settings.workbook.styles[style_name]=url
    }
  
    if(jade_settings.workbook.styles[style_name].substr(0,8)==="https://"){
      // the style has not yet been fetched
      //console.log("fetching")
      const response = await fetch(jade_settings.workbook.styles[style_name])
      const data = await response.text()
      //console.log("data",data)
      jade_settings.workbook.styles[style_name]=data
      if(integrate_now){
        document.getElementById("head_style").remove()
        document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + jade_settings.workbook.styles[style_name] + "</style>")
      }
    }
    
  }
  static set_style(style_name){

   //console.log("at set style", style_name)
   //console.log("jade_css_suffix", jade_css_suffix)
    //let css_sfx = jade_css_suffix
    if(!style_name){
      style_name="system"
      //css_sfx=""
  
    }
  
    if(jade_settings.workbook.styles[style_name].substr(0,8)==="https://"){
      // this style has not been fetched.  Get it now
      //console.log("in iff")
      Jade.get_style(style_name)
    }
  
    const style_tag = document.getElementById("head_style")
    if(style_tag.dataset.name!==style_name ){
      // only update the style tag if it is a differnt name
      style_tag.remove()
      document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="' + jade_settings.workbook.styles[style_name] + '">' + the_style + "</style>")
    }
    const user_style_tag = document.getElementById("user_style")
    if(user_style_tag ){
      user_style_tag.remove()
    }
    document.head.insertAdjacentHTML("beforeend", '<style id="user_style">' + jade_css_suffix + "</style>")

    

    //set the user style

    
  }
  static panel_close_button(panel_name){
    return '<div id="close_' + panel_name + '" onclick="Jade.close_canvas(\'' + panel_name + '\')" class="top-corner" style="padding:5px 5px 0 5px;margin:5px 15px 0 0; cursor:pointer"><i class="far fa-window-close fa-2x"></i></i></div>'
  }
  static build_panel(panel_name, show_close_button){
    const div=document.createElement("div");
    div.id=panel_name
    div.style.display="none"
    div.innerHTML=Jade.panel_close_button(panel_name)
    if(!jade_panels.includes(panel_name)){
      jade_panels.push(panel_name)
    }
    document.body.appendChild(div)
    if (show_close_button===undefined){
      show_close_button=true
    }
    if(!show_close_button){
      Jade.hide_element("close_" + panel_name)
    }
  }
  static show_automations(show_close_button, heading, list){
    //console.log("at list automations")
    const panel_name="panel_listings"
    // get the list of functions
    //###################################################### need to iterate over all modules
    const html=[`<h2 style="margin:0 0 0 1rem">${heading || 'Active Automations'}</h2><ol>`]
    //console.log("code panels", jade_code_panels)
    if(list){// a list has been provide, use it
      //console.log("list", list)
      for(const entry of list){
        html.push('<li onclick="'+entry.action+'" style="cursor:pointer"><b>'+entry.name+'</b>: '+entry.description+'</li>')
      }
    }else{// no list has been provide, so we read from the code modules in this workbook
      for(const code_panel of jade_code_panels){
        const editor = ace.edit(code_panel + "-content");
        const code = editor.getValue();
        const parsed_code=Jade.parse_code(code)
    
    
        //console.log(parsed_code)
        
        if(!parsed_code.error){ 
          for(const element of parsed_code.body){
            let call_stmt = null
            if(element.type==="FunctionDeclaration"){
              if(element.id && element.id.name){
                // this is a named function
                //console.log("found", element)
                if(element.params.length===0){
                  //there are no params. it is callable
                  call_stmt=element.id.name + "()"
                } else if(element.params.length===1){
                  // there is one param.  
                  if(element.async){
                    if(("excel ctx context").includes(element.params[0].name)){
                      // this is an async function with a single parameter named excel, ctx or context.  Run by passing context
                      call_stmt = `Excel.run(jade_modules.${Jade.panel_name_to_module_name(code_panel)}.${element.id.name})`
                    }
                  }
                }  
                if(call_stmt){ // this is a function we can run directly
                  // check for comment
                  const function_text = jade_modules[Jade.panel_name_to_module_name(code_panel)][element.id.name]+''
                  //const function_text = window[element.id.name]+''
                  //console.log(function_text)
      
      
                  if(function_text.toLowerCase().includes("jade.listing:")){ // this is a function we can run directly and it as the comment
                    const start = function_text.toLowerCase().indexOf("jade.listing:")+13
                    const num_chars =  function_text.toLowerCase().split("jade.listing:")[1].split("*/")[0].length
                    const comment = function_text.substring(start, start+num_chars)
                    try{
                      const comment_json=JSON.parse(comment)
                      html.push('<li onclick="'+call_stmt+'" style="cursor:pointer"><b>'+comment_json.name+'</b>: '+comment_json.description+'</li>')
                    }catch(e){
                      ;console.log("Jade.listing was not valid JSON", comment)        
                    }
                  }//for function on code page
                } 
              }
            }
          } 
        }
      }
    }
  
    if(html.length===1){
      //Did not find any properly configured functions
      html.push("There are currently no active automations in this workbook.")
    }else{
      html.push("</ul>")  
    }
    Jade.open_canvas("panel_listings",html.join(""), show_close_button)
  }
  static show_panel(panel_name){

    if(jade_code_panels.includes(panel_name)){
      // set the size in case it is off
      if(tag(panel_name + "_function-names").length===0){
        // there are no function to run
        Jade.hide_element(panel_name + "_function-names")
      }
      tag(panel_name + "_editor-page").style.height = Jade.editor_height()
      try{
        ace.edit(panel_name + "-content").focus()
      }catch(e){
        //console.log("could not access ace.  This is expected",e)
      }
    }
    //################## 3 is  a problem
    if(jade_panels.slice(0, 3).includes(panel_name) || jade_code_panels.includes(panel_name)){
      Jade.set_style()
    }
    
   //console.log("trying",panel_name)
    for(const panel of jade_panels){
      if(panel===panel_name){
       //console.log("showing", panel)
        if(tag("selector_" + panel_name)){
          tag("selector_" + panel_name).value=panel_name
        }
        tag(panel).style.display="block"  
        tag(panel).style.height="100%"  
        jade_panel_stack.push(panel)
      }else{
        //console.log(" hiding", panel)
        if(panel==="panel_examples"){
          if(tag(panel)?.style.display!=="none"){  
            // close and destroy examples interface so we don't have conflicts in ID if hte example gets used
            for(const elem of document.getElementsByClassName("e-canvas")){
              // clear all example canvases
              elem.innerHTML=""
            }
            for(const elem of document.getElementsByClassName("e-page")){
              // hide all editors
              elem.style.display = "none"
              
            }
            
          }
        }
        if(tag(panel)){
          tag(panel).style.display="none"
        }
      }
    }
  
    if(jade_code_panels.includes(panel_name)){
      //focus the ace editor
      try{
        ace.edit(panel_name + "-content").focus()
      }catch(e){
        //;console.log("could not access ace.  This is expected",e)
      }
    }
  
  }
  static toggle_wrap(panel_name){
    const editor = ace.edit(panel_name + "-content");
    if(editor.getOptions().wrap==="off"){
      editor.setOptions({wrap: true})
      globalThis=true
    }else{
      editor.setOptions({wrap: "off"})
    }
  }
  static add_code_editor(module_name, code, module_xmlid, mod_settings, options_in){
    // jade_settings are things gove is storing with the module
    // options are the options from the ace editor
   //console.log("jade_settings", jade_settings)
    let options = jade_settings.user.ace_options
  
    
    // not currently handling options at the editor level, so this block is diabled
    // if(options_in){// default options for the editor
    //   options=options_in
    // }
  
    //console.log(module_name, "options", options)
    if(!mod_settings){
      mod_settings={cursorPosition:{row: 0, column: 0}}
    }
    
    //console.log("adding ace editor", module_name, module_xmlid)
    const panel_name = "panel_" + module_name.toLowerCase().split(" ").join("_") + "_module"
    jade_code_panels.push(panel_name)
    jade_panel_labels.push(Jade.panel_name_to_panel_label(panel_name))
    jade_panels.push(panel_name)
    Jade.build_panel(panel_name, false)
    tag(panel_name).dataset.module_name = module_name
    tag(panel_name).dataset.module_xmlid = module_xmlid
    
   //console.log(2)
  
   //console.log("initializing examples", tag(panel_name))
    
    tag(panel_name).appendChild(Jade.get_panel_selector(panel_name))
    const editor_container = document.createElement("div")
    editor_container.className=panel_name+"-content"
  
    let div = document.createElement("div");
    div.style.padding=".2rem"
    div.style.verticalAlign="middle"
    div.innerHTML = '<button title="Save code to workbook" onclick="Jade.update_editor_script(\'' + panel_name + '\')">Save</button> <button onclick="Jade.code_runner(tag(\'' + panel_name + '_function-names' + '\').value,\'' + panel_name + '\')" title="Save code to workbook and execute">Run</button> <select id="' + panel_name + '_function-names"></select>';
    div.style.height="22px"
    div.style.fontFamily="auto";
    div.style.fontSize = "1rem";
    div.style.padding=".2rem"
    div.style.backgroundColor = "#eee";
    div.id=panel_name + "_editor-bar"
    editor_container.appendChild(div);
   //console.log(3)  onclick="Jade.code_runner(tag(\'' + panel_name + '_function-names' + '\').value,\'' + panel_name + '\')"
  
  
    //console.log("=======================================")
    //console.log(div);
    //console.log(div.clientHeight);
  
    const box = document.createElement("div");
    box.id = panel_name + "_editor-page";
    box.style.width = "100%";
    box.style.height = Jade.editor_height()
    box.style.display = "inline-block";
    box.style.position = "relative";
   //console.log(4)
  
   //console.log("document",document.body.clientHeight);
   //console.log("scr",tag("panel_code_editor").Height);
  
  
    div = document.createElement("div");
    div.id = panel_name + "-content";
    div.dataset.edited=false
    div.innerHTML = code.toHtmlEntities();
  
    box.appendChild(div);
   //console.log(5)
  
  
    editor_container.appendChild(box);
  
    //        elem.innerHTML = '<pre id="pre' + id + '">' + gist.script.content.split("<").join("&lt;") + '</pre>'
  
    const scriptPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      document.body.appendChild(script);
      script.onload = resolve;
      script.onerror = reject;
      script.async = true;
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.13/ace.js";
    });
   //console.log(6)
  
    scriptPromise.then(() => {
      const editor = ace.edit(panel_name + "-content");
      
      editor.on("blur", function () {
        window.event.target.parentElement.dataset.edited=true
      });
      //editor.setTheme("ace/theme/solarized_light");
      // if(!isNaN(options.fontSize)){
      //   options.fontSize += "pt"
      // }
      editor.setOptions(options);
      editor.session.setMode("ace/mode/javascript");
  
      //console.log("jade_settings", jade_settings)
      editor.moveCursorTo(mod_settings.cursorPosition.row, mod_settings.cursorPosition.column)
  
      editor.commands.addCommand({  // toggle word wrap
        name: "wrap",
        bindKey: {win: "Alt-z", mac: "Alt-z"},
        exec: function(editor) {
          for(const panel of jade_code_panels){
            if(tag(panel).style.display==="block"){
              // we found the one that is visible
              Jade.toggle_wrap(panel)
              break//exit the loop
            }
          }
        }
      })
  
      editor.commands.addCommand({  // could do ctrl+r but want to be parallel with save
        name: "run",
        bindKey: {win: "Ctrl-enter", mac: "Command-enter"},
        exec: function(editor) {
          for(const panel of jade_code_panels){
            if(tag(panel).style.display==="block"){
              // we found the one that is visible
              Jade.code_runner(tag(panel + '_function-names').value, panel)
              break//exit the loop
            }
          }
        }
      })
  
      editor.commands.addCommand({  // could do ctrl+r but want to be parallel with save
        name: "run_shift",
        bindKey: {win: "alt-enter", mac: "alt-enter"},
        exec: function(editor) {
          for(const panel of jade_code_panels){
            if(tag(panel).style.display==="block"){
              // we found the one that is visible
              Jade.code_runner(tag(panel + '_function-names').value, panel)
              break//exit the loop
            }
          }
        }
      })

      //Now that we have loaded this code editor, try to run it's autoec
      if(jade_modules[Jade.panel_name_to_module_name(panel_name)].auto_exec){
         // if this module has an autoexec, let's run it
         jade_modules[Jade.panel_name_to_module_name(panel_name)].auto_exec()
      }
    });
  
    
  
    editor_container.style.display = "block";
  
    const parsed_code=Jade.parse_code(code)
    if(!parsed_code.error){
      // only incorporate the code if it is free of syntax errors
      Jade.incorporate_code(code, module_name)  
    }
  
    
    
    tag(panel_name).appendChild(editor_container)
  
    Jade.load_function_names_select(code, panel_name)
    if(mod_settings.func){
      tag(panel_name + "_function-names").value=mod_settings.func
  
    }
 
  // now running auto exec when teh codeitor loads  
  //   // AutoExecutable function
  // //console.log("about to autoexec")
  //   try{
  //   //console.log("in try", Jade.panel_name_to_module_name(panel_name))
  //     //jade_modules[Jade.panel_name_to_module_name(panel_name)].auto_exec()
  //    //console.log("after autoexec", jade_modules, JSON.stringify(jade_modules))
  //     jade_modules.code.auto_exec()
  //   }catch(e){
  //    ;console.log("catch",e)
  //   }  



  }
  static code_runner(script_name,panel_name){
    //console.log('panel_name',panel_name)
    //console.log('script_name',script_name)
    if (tag(panel_name + "-content").dataset.edited==="true"){
      if(!Jade.update_editor_script(panel_name)){
        // Jade.update_editor_script returns false if there is a 
        // syntax error.  Don't run the old code
        return
      }
    }
    //console.log("panel_name", panel_name, this.panel_name_to_module_name(panel_name))
    
    if(script_name.includes("(excel)")){
      Excel.run(jade_modules[Jade.panel_name_to_module_name(panel_name)][script_name.split("(")[0]])
    }else{
      //console.log("panel_name",panel_name)
      //console.log("Jade.panel_name_to_module_name(panel_name)",Jade.panel_name_to_module_name(panel_name))
      const mod_name=Jade.panel_name_to_module_name(panel_name)
      // //console.log("mod_name",mod_name, script_name)
      // //console.log("jade_modules[mod_name][script_name]",jade_modules[mod_name][script_name])
      jade_modules[mod_name][script_name]()
    }
  }
  static init_output(){
    const panel_name="panel_output"
    Jade.build_panel(panel_name, false)
    const panel=tag(panel_name)
   //console.log("initializing examples")
    panel.appendChild(Jade.get_panel_selector(panel_name))
    Jade.print('This panel shows the results of your calls to the "print" function.  Use "Jade.print(data)" to append text to the most recently printed block.', "About the Output Panel")
    Jade.print('\nUse "Jade.print(data, heading)" to start a new block.')
    
  }
  static rebuild_examples(gist_id){
    tag("panel_examples").innerHTML=""
    Jade.fill_examples(gist_id)
  }
  static init_examples(){
    Jade.build_panel("panel_examples", false)
    //Jade.fill_examples()  // now it is filled when the examples are first shown.
  }  
  static fill_examples(gist_ids){
    //gist_ids is a comma delimited list of gist ids that hold examples
    
    if(!gist_ids){gist_ids=jade_settings.workbook.examples_gist_id}
    const panel_name="panel_examples"
    const panel=tag(panel_name)
   //console.log("initializing examples")
    panel.appendChild(Jade.get_panel_selector(panel_name))
    const div = document.createElement("div")
    div.className="content"
    div.id="e_content"
    panel.appendChild(div)
   //console.log("gist_ids",gist_ids, jade_settings)
  
    const gist_list=[]
   //console.log("gist_ids.split(",")",gist_ids.split(","))
    for(const gist of gist_ids.split(",")){
     //console.log("gisting",gist)
      gist_list.push(gist.trim())
    }
   //console.log("gist_list",gist_list)
    const html=[]
    for(const gist of gist_list){
      html.push(`<div id="${gist}"></div>`)
    }
    tag("e_content").innerHTML = html.join("");
  
    for(let i=0;i<gist_list.length;i++){
     //console.log("gist_list[i]",gist_list[i])
      Jade.get_example_html(gist_list[i],i+1)// get the gist and integrate the examples
    }
  }
  static get_example_html(gist_id, sequence){
    
    let url=`https://api.github.com/gists/${gist_id}?${Date.now()}`
    if(gist_id="f8e5fc20cff0c19a27765e7ce5fe54fe"){
      // the default gist, read it from jade-addin's gist server
      // we do this because the examples get loaded very often.
      
      url=`https://jade-addin.github.io/public-gist-server/${gist_id}`

    }
    const gist_url=`https://gist.github.com/${gist_id}`
   //console.log("building examples",url)
    
   //console.log("about to fetch",url)
    fetch(url)
    .then((response) => response.text())
    .then((json_text) => {
   //console.log("json_text",json_text)
      const data=JSON.parse(json_text)
  
      // now we have the data from the api call.  need to organize it--especially for order
      // make a list of objects with integers as keys so the order will be numerical
      const filenames={}
      for(const filename of Object.keys(data.files)){
        filenames[parseFloat(filename.split("_").shift())] = filename
      }
      //console.log("filenames",filenames)
      let temp = data.description.split(":")
      const html=[`<h2 style="cursor:pointer" title="Copy link to Example" onclick="Jade.copy_to_clipboard('${gist_url}')">${temp.shift()}</h2>`]
      html.push(temp.join(":").trim())
      // iterate for each file in the gist
      let example_number=0
      for(const key of Object.keys(filenames)){
        example_number++
        //console.log(key, filenames[key])
        temp=filenames[key].split(".")
        temp.pop()//remove the extension
        temp=temp.join(".").split("_")
        temp.shift()// remove the sort sequence number
        html.push(`<p><a class="link" title="Show Code" onclick="Jade.show_example('${sequence*100+example_number}','${data.files[filenames[key]].raw_url}',${data.files[filenames[key]].content.split(/\r\n|\r|\n/).length})"><b>${example_number}. ${temp.join(" ")}:</b> </a>`);
        temp=data.files[filenames[key]].content

        if(temp.includes("*/")){
          // there is a block comment.  assume it is a description
          html.push(temp.split("*/")[0].split("/*")[1])  
        }
        html.push(`</p><div class="e-page hidden" id="page${sequence*100+example_number}"></div>`)
        //${data.files[filenames[key]].content}//eeeeeeeeeeeeeeeeeee
      }
      tag(gist_id).innerHTML = html.join("");
      example_number=0
      for(const key of Object.keys(filenames)){
        example_number++
        // now place the example html.  Originally, this code was written to bring in the html again when clicked
        Jade.place_example(sequence*100+example_number,data.files[filenames[key]].content,data.files[filenames[key]].content.split(/\r\n|\r|\n/).length)
      }
      return
  
    });
  
  
  }



  static place_example(id,data,lines) {
    const elem = tag("page" + id);
   //console.log("elem",elem)
    let div = document.createElement("div");
    div.id="example" + id + "_html"
    div.className = "e-canvas";
    div.style.marginBottom = "1rem";
    elem.appendChild(div);
    let box_height = (lines*22+17)// the size neede to show the whole example
    if(box_height>window.innerHeight){
      box_height=window.innerHeight-60
    }
    
    const box = document.createElement("div");
    box.id = "page_" + id;
    box.style.width = "100%";
    box.style.height = box_height+"px";
    box.style.display = "inline-block";
    box.style.position = "relative";


    div = document.createElement("div");
    div.id = "editor" + id;
    div.innerHTML = data.toHtmlEntities()

    box.appendChild(div);
    elem.appendChild(box);
    
    
    // setting up the editor in the example space
    const scriptPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      document.body.appendChild(script);
      script.onload = resolve;
      script.onerror = reject;
      script.async = true;
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.13/ace.js";
    });

    scriptPromise.then(() => {
      const editor = ace.edit("editor" + id);
      editor.on("blur", function () {
        Jade.update_script(id);
      });
      editor.setTheme("ace/theme/tomorrow");
      //editor.session.$worker.send("changeOptions", [{ asi: true }]);
      editor.setOptions({fontSize: "14pt"});
      editor.getSession().setUseWorker(false);
      editor.session.setMode("ace/mode/javascript");
      editor.commands.addCommand({  // toggle word wrap
        name: "wrap",
        bindKey: {win: "Alt-z", mac: "Alt-z"},
        exec: function(editor) {            
          if(editor.getOptions().wrap==="off"){
            editor.setOptions({wrap: true})
            globalThis=true
          }else{
            editor.setOptions({wrap: "off"})
          }              
        }
      })
    })

    

    Jade.incorporate_code(data + "\n" + Jade.show_example_html_script(id),"jade_examples")
    //jade_modules.add(setup)
    jade_modules.jade_examples.setup() // setup must be defined in the example
    if(!Jade.is_visible(tag("page_" + id ))){
      tag("page_" + id ).scrollIntoView(false)
    }
    return elem

  }


  
  static show_example(id, url, lines) {
    // place the code in an editable box for user to see and play with
    // these examples should be made in script lab to have the right format

    const elem = tag("page" + id);
    //console.log(id + id);
    if (elem.innerHTML === "") {
     //console.log("lines", lines)
      elem.innerHTML='<img id="loading-image" width="50" src="assets/loading.gif" />'
  
      fetch(url)
        .then((response) => response.text())
        .then((data) => {
         //console.log(data)
          //const gist = jsyaml.load(data);
          //console.log(gist)
          tag("loading-image").remove()
          const elem=this.place_example(id,data,lines) 
          elem.style.display = "block";
        })
        .catch((error) => {
         ;console.error(error);
        })
    } else {
      if (elem.style.display === "block") {
        elem.style.display = "none";
      } else {
        elem.style.display = "block";
        Jade.incorporate_code(ace.edit("editor" + id).getValue() + "\n" + Jade.show_example_html_script(id),"jade_examples")
        //jade_modules.add(setup)
        jade_modules.jade_examples.setup() // setup must be defined in the example
        if(!Jade.is_visible(tag("page_" + id ))){
          tag("page_" + id ).scrollIntoView(false)
        }
    }
    }
    
  }
  static show_example_html_script(id){
    return 'function show_html(html){tag("example'+id+'_html").innerHTML=html}\n'
  }
  static copy_to_clipboard(text) {
    navigator.clipboard
      .writeText(text)
  }
  static update_script(id) {
    // read the script for an ace editor and write it to the DOM
    // this is the one used by the examples page
    //console.log("script" + id);
    const editor=ace.edit("editor" + id)
    const code = editor.getValue()
   //console.log(code)
    const parsed_code=Jade.parse_code(code)
  
    if(parsed_code.error){
      alert(parsed_code.error, "Syntax Error")
      editor.moveCursorTo(parsed_code.error.split("(")[1].split(":")[0]-1,parsed_code.error.split("(")[1].split(":")[1].split(")")[0])
      editor.focus()
      return false
    }
   //console.log(1)
    Jade.incorporate_code(code + "\n" +Jade.show_example_html_script(id), "jade_examples")
    
   //console.log(2)

    
    return true
  }    

  static parse_code(code){
    try{
      return acorn.parse(code, { "ecmaVersion": "latest" }  );
    }catch(e){
      return {error:e.message}
    }
  }

  static update_editor_script(panel_name) {
    // read the script for an ace editor and write it to the DOM
    // also saves the module to the custom properties
    //console.log("at Jade.update_editor_script", panel_name)
    // set the size of the editor in case there was a prior zoom
  
    // get the code
    const editor=ace.edit(panel_name + "-content")
    const code = editor.getValue();
  
    // save the script to the workbook.  This is the most important thing
    // we are doing at the moment.  Do it first
   //console.log("about to write module")
    Jade.write_module_to_workbook(code, panel_name)
  
    // update the height of the editor in case it has gotten out of synch
    tag(panel_name + "_editor-page").style.height = Jade.editor_height()
  
  
    // show_html is defined differntly for modules than for examples
    // we need to be sure it is defined correctly for modules, so we set it here
  //Jade.incorporate_code('function show_html(html){open_canvas("html", html)}')
  
  
    //Check for syntax errors
    const parsed_code=Jade.parse_code(code)
  
    if(parsed_code.error){
      alert(parsed_code.error, "Syntax Error")
      editor.moveCursorTo(parsed_code.error.split("(")[1].split(":")[0]-1,parsed_code.error.split("(")[1].split(":")[1].split(")")[0])
      editor.focus()
  
      return false
    }
    
    // load the user's code into the browser
    Jade.incorporate_code(code+ '\nfunction show_html(html){Jade.open_canvas("html", html)}' ,Jade.panel_name_to_module_name(panel_name))

    // put the function names in the function dropdown
    Jade.load_function_names_select(parsed_code, panel_name)
    
    return true
  }
  static incorporate_code(code, module_name_in){
    // It turns out that the script block does not need to persist in the HTML
    // once the script block is loaded, the JS is parsed and not again referenced.
    // So, we create a script block, append it to the document body, then remove.
    //console.log(module_name_in, code)
    let new_code=code
    if(module_name_in){
      // a module_name is supplied, put this in Jade Modules
      const module_name=module_name_in.toLowerCase()
      const parsed_code = Jade.parse_code(code)
      if(parsed_code.error){
        ;console.error("Error in",module_name_in, parsed_code)
        return
      }
      const function_names=[]
      for(const element of parsed_code.body){
        if(element.type==="FunctionDeclaration"){
          if(element.id && element.id.name){
          //console.log("Found:",element.id.name)
            function_names.push(element.id.name)
          }
        }else if(element.type==="VariableDeclaration" && element.declarations[0] && element.declarations[0].init&&element.declarations[0].init.type.includes("FunctionExpression")){
          if(element.declarations[0].id && element.declarations[0].id.name){
            // this is a named function
          //console.log("Found:",element.declarations[0].id.name)
            function_names.push(element.declarations[0].id.name)
          }
        }
      } 
      //console.log("module_name",module_name)
      //console.log("function_names",function_names)
      //console.log("=====================================================")
      //console.log("====================================================")
      //console.log("=====================================================")
      //console.log(code)
      //console.log("=====================================================")
      //console.log("====================================================")
      //console.log("=====================================================")


      new_code = `var JADE_USER_CODE_MASTER = function (){${code}
        ;jade_modules.add('${module_name}',${function_names})
      }();`
    }
     
    const script = document.createElement("script");
    script.innerHTML = new_code
    document.body.appendChild(script);
    document.body.lastChild.remove()
  }


  static write_module_to_workbook(code, panel_name){
    let options = {fontSize:"14pt"}
    const settings = {
      cursorPosition:{row: 0, column: 0},
      func:tag(panel_name + "_function-names").value
    }
  
    try{
      const editor = ace.edit(panel_name + "-content") 
      jade_settings.cursorPosition = editor.getCursorPosition()
      options=editor.getOptions()
    }catch(e){
      //console.log("This is an expected error: ace editor not yet built",e)
    }
  
    //console.log("writing options", JSON.stringify(options))
    
    const module_name = tag(panel_name).dataset.module_name
    Jade.save_module_to_workbook(code, module_name, options)
  }
  static async save_module_to_workbook(code, module_name, options){
    const foundId = jade_settings.workbook.code_module_ids.find(id=>{
      return Jade.getSetting(`${id}-module-name`) === module_name
    })

    if(foundId){
      Jade.setSetting(`${foundId}-module-code`, btoa(code))
      Jade.setSetting(`${foundId}-module-options`, JSON.stringify(options))
    }else{
      const id = uuid()
      jade_settings.workbook.code_module_ids.push(id)
      Jade.write_settings_to_workbook()
      Jade.setSetting(`${id}-module-code`, btoa(code))
      Jade.setSetting(`${id}-module-name`, module_name)
      Jade.setSetting(`${id}-module-options`, JSON.stringify(options))
    }
  }

  static get_module(id){
    console.log("get_module", id)
    return {
      code: Jade.getSetting(`${id}-module-code`),
      name: Jade.getSetting(`${id}-module-name`),
      options: Jade.getSetting(`${id}-module-options`),
    }
  }

  static load_function_names_select(code,panel_name){// reads the function names from the code and puts them in the function name select

    // if a string is passed in, parse it.  otherwise, assume it is already parsed
    if(typeof code === "string"){
      var parsed_code = Jade.parse_code(code)
    }else{
      var parsed_code = code
    }
    
    if(parsed_code.error){
      return
    }
  
    const selectElement=tag(panel_name + "_function-names")
    const selected_script = selectElement.value
    
  
    while(selectElement.options.length>0) {
       selectElement.remove(0);
    }
  

    for(const element of parsed_code.body){
      let option_value = null
      if(element.type==="FunctionDeclaration"){
        if(element.id && element.id.name){
          // this is a named function
          if(element.params.length===0){
            //there are no params. it is callable
            option_value=element.id.name// + "()"
          } else if(element.params.length===1){
            // there is one param.  
            if(element.async){
              if(("excel ctx context").includes(element.params[0].name)){
                // this is an async function with a single parameter named excel, ctx or context.  Run by passing context
                option_value = element.id.name + "(excel)"
              }
            }
          }  
          if(option_value){ // this is a function we can run directly
            const option = document.createElement("option")
            if(option_value===selected_script){option.selected='selected'}
            option.text = element.id.name
            option.value = option_value
            selectElement.add(option)    
          } 
        }
      }else if(element.type==="VariableDeclaration" && element.declarations[0] && element.declarations[0].init&&element.declarations[0].init.type.includes("FunctionExpression")){
        if(element.declarations[0].id && element.declarations[0].id.name){
          // this is a named function
          if(element.declarations[0].init.params.length===0){
            //there are no params. it is callable
            option_value=element.declarations[0].id.name// + "()"
          } else if(element.declarations[0].init.params.length===1){
            // there is one param.  
            if(element.declarations[0].init.params.length.async){
              if(("excel ctx context").includes(element.declarations[0].init.params[0].name)){
                // this is an async function with a single parameter named excel, ctx or context.  Run by passing context
                option_value = element.declarations[0].id.name + "(excel)"
              }
            }
          }  
          if(option_value){ // this is a function we can run directly
            const option = document.createElement("option")
            if(option_value===selected_script){option.selected='selected'}
            option.text = element.declarations[0].id.name
            option.value = option_value
            selectElement.add(option)    
          } 
        }
      }
    } 
    if(selectElement.length===0){
      Jade.hide_element(selectElement)
    }else{
      Jade.show_element(selectElement)
    }
  }
  static get_panel_selector(panel){
    const panel_label=Jade.panel_name_to_panel_label(panel)
    const sel = document.createElement("select")
    //console.log("appending panel=====", panel)
    if(!jade_panel_labels.includes(panel_label)){
      jade_panel_labels.push(panel_label)
    }
    sel.className="panel-selector"
    // put the options in this panel selector
    Jade.update_panel_selector(sel)
  
    // update all others panel selectirs
    for(const selector of document.getElementsByClassName("panel-selector")){
      Jade.update_panel_selector(selector)
    }
    
    sel.value=panel
    sel.style.height="40px"
    sel.id="selector_"+panel
    sel.onchange = Jade.select_page
    return sel
  }
  static update_panel_selector(sel){
    // put the proper choices in a panel selector
    while(sel.length>0){
       sel.remove(0)
    }
    for (let i=0; i<jade_panel_labels.length; i++) {
      var option = document.createElement("option");
      option.value = Jade.panel_label_to_panel_name(jade_panel_labels[i]) 
     //console.log("-->", option.value)
      option.text = jade_panel_labels[i];
      option.className="panel-selector-option"
      sel.appendChild(option);
    }
  }
  static panel_label_to_panel_name(panel_label){
    return "panel_" + panel_label.toLocaleLowerCase().split(" ").join("_");
  }
  static panel_name_to_panel_label(panel_name){
    let panel_label = panel_name.replace("panel_","")
    panel_label=panel_label.split("_").join(" ") 
    return panel_label.toTitleCase()
  }
  static panel_name_to_module_name(panel_name){
    // this is the name used to refer to modules from other modules
    const data=panel_name.substr(6, panel_name.length-13)
    //console.log("data in panel_name_to_module_name",data)
    return data
  }
  static file_name_to_module_name(file_name){
    // this is the name used to refer to modules from other modules
    //console.log("a")

    let new_name=file_name.toLowerCase()
    //console.log("b")
    if(new_name.slice(-3)===".js"){
      //console.log(2)
      new_name = new_name.slice(0,new_name.length-3)
    }
    //console.log(3)
    new_name=new_name.split(" ").join("_")
    //console.log(4)
    new_name=new_name.split(".").join("_")
    //console.log(5)
    return new_name
  }
  static select_page(){
   //console.log("panel name",window.event.target.value)
    Jade.show_panel(window.event.target.value)
  }
  static show_element(tag_id){
    // removes the hidden class from a tag's css
    if (typeof tag_id==="string"){
      var the_tag=tag(tag_id)
    }else{  
      the_tag=tag_id
    }
    the_tag.className=the_tag.className.split("hidden").join("")
  }  
  static hide_element(tag_id){
    // adds the hidden class from a tag's css
    //console.log("tag_id",tag_id)
    //console.log("tag(tag_id)",tag(tag_id))
    //console.log("tag(tag_id).className",tag(tag_id).className)
    if (typeof tag_id==="string"){
      var the_tag=tag(tag_id)
    }else{  
      the_tag=tag_id
    }
    if(the_tag){
      if(the_tag.className){
        if(!the_tag.className.includes("hidden")){
          the_tag.className=(the_tag.className + " hidden").trim()
        }
      }else{
        the_tag.className="hidden"
      }
    }
  }
  static toggle_element(tag_id){
    // adds the hidden class from a tag's css
   //console.log("tag_id", tag_id)
    if (typeof tag_id==="string"){
      var the_tag=tag(tag_id)
    }else{  
      the_tag=tag_id
    }
  
    if(the_tag.className.includes("hidden")){
      Jade.show_element(the_tag)
    }else{
      Jade.hide_element(the_tag)
    }
  }
  static editor_height(){
    return (window.innerHeight-73)+"px"
  }
  static default_code(panel_name){
    //console.log("platform:",PLATFORM)
    if(PLATFORM==="PowerPoint"){
      return powerpoint_module()
    }else if(PLATFORM==="Word"){
      return word_module()
    }
    let code = `async function write_timestamp(excel){
  /*Jade.listing:{"name":"Timestamp","description":"This sample function records the current time in the selected cells"}*/
  excel.workbook.getSelectedRange().values = new Date();
  await excel.sync();
}
  
function auto_exec(){
  // This function is called when the addin opens.
  // un-comment a line below to take action on open.
  
  // Jade.open_automations() // displays a list of functions for a user
`
    if(panel_name){
      code += `  // Jade.show_panel('${panel_name}')      // shows this code editor
}`
    }else{
      code += `  // Jade.open_editor()      // shows the code editor
}`
    }
    return code
  }


  static show_import_module(){
   //console.log("at show import", jade_settings.workbook.module_to_import)
    if(jade_settings.workbook.module_to_import){
     //console.log("in if")
      tag("gist-url").value=jade_settings.workbook.module_to_import
    }
    Jade.toggle_element('import-module')
    tag('gist-url').focus()
  
  }
  static is_visible(el) {
    const rect = el.getBoundingClientRect();
    return (
        rect.top >= 0 &&
        rect.left >= 0 &&
        rect.bottom <= (window.innerHeight || document.documentElement.clientHeight) &&
        rect.right <= (window.innerWidth || document.documentElement.clientWidth)

    );
}

static getSetting(key){
  return Office.context.document.settings.get(key)
}

static setSetting(key, value){
  Office.context.document.settings.set(key, value);
  return Jade.saveSettingsToDocument()
}

static async saveSettingsToDocument(){
  return await new Promise(resolve => {
    Office.context.document.settings.saveAsync(resolve);
  })
}

}// end of Jade class


/*** Convert a string to HTML entities */
String.prototype.toHtmlEntities = function() {
  return this.replace(/./gm, function(s) {
      // return "&#" + s.charCodeAt(0) + ";";
      return (s.match(/[a-z0-9\s]+/i)) ? s : "&#" + s.charCodeAt(0) + ";";
  });
};


String.prototype.toTitleCase=function() {
  let str = this.toLowerCase().split(' ');
  for (var i = 0; i < str.length; i++) {
    str[i] = str[i].charAt(0).toUpperCase() + str[i].slice(1); 
  }
  return str.join(' ');
}

function alert(data, heading, headerColor){
  if(tag("jade-alert")){tag("jade-alert").remove()}
  if(!heading){heading="System Message"}
  const div = document.createElement("div")
  div.className="jade-alert"
  div.id='jade-alert'
  const header = document.createElement("div")
  if(headerColor){
    header.style.backgroundColor = headerColor
  }
  header.className="jade-alert-header"  
  header.innerHTML = heading 

  const close_button=document.createElement("div")
  close_button.className="jade-alert-close"
  close_button.innerHTML='<i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.remove()">'
  const body = document.createElement("div")
  body.className="jade-alert-body"  
  body.innerHTML = data
  div.appendChild(header)
  div.appendChild(close_button)
  div.appendChild(body)
  document.body.appendChild(div)

}
function tag(id){
  // a short way to get an element by ID
  return document.getElementById(id)
}

function powerpoint_module(){
  // default code module when running under powerpoint
  return `function run_powerpoint(){
    Jade.load_js('http://localhost:5500/powerpoint.js')
  }

  function load(){
    Jade.load_js('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js')
  }`
}
function word_module(){
    // default code module when running under word
    return `function run_word(){
      Jade.load_js('http://localhost:5500/word.js')
    }`
  
}

function uuid() {
  return "10000000-1000-4000-8000-100000000000".replace(/[018]/g, c =>
    (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
  );
}
    


class jade extends Jade{}// just let lowercase call to jade work
class JADE extends Jade{}// just let uppercase call to jade work