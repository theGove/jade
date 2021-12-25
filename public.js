//  these are the functions that are documented for use by users of the addin
const global={}
const default_style_name="marx"

// We can use gist as repository for the code then either
// consume it or import it.  Gists have multiple files
// that become modules in jsvba

async function load_gist(gist_id){
    console.log("gisting", gist_id)
  try{
      
    const response = await fetch(`https://api.github.com/gists/${gist_id}?${new Date()}`)
    const data = await response.json()
    for(const file of Object.values(data.files)){
        console.log("===================================")
        console.log(file.content)
        console.log("===================================")
        incorporate_code(file.content)
    }
    auto_exec()
    incorporate_code("auto_exec=null")
    
  }catch(e){
    console.log("Error fetching gist", e)
  }
}


async function load_code_module(gist_url){
    //takes a gist URL and loads it's content in a script window
    const code=await get_code_from_gist(gist_url)
    const script = document.createElement("script")
    script.innerHTML = code
    script.id="gist-import-script"
    document.body.appendChild(script)
    tag("gist-import-script").remove()
  
}

async function import_code_module(url_or_gist_id){
    settings.workbook.module_to_import=url_or_gist_id
    console.log("at import code mod", settings.workbook)

    hide_element("import-module")
    Jade.save_settings()
    if(!url_or_gist_id){
        return
    }
    let url=null
    if(url_or_gist_id.substr(0,4)==="http"){
        // check to see if GIST url
        if(url_or_gist_id.toLowerCase().includes("gist.github.com")){
            const url_data = url_or_gist_id.split("/")
            url='https://api.github.com/gists/' + data[data.length-1] + "?" + new Date()
        }else{
            url=url_or_gist_id
        }

    }else{
        // this looks like a gist id.  we should probably check it
        // sometime.  for now, let's just assume it is
        url = 'https://api.github.com/gists/' + url_or_gist_id + "?" + new Date()
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
            console.log("parsing json from gist")
            for(const file of Object.values(gist_json.files)){
                files.push({name:file.filename.split(".js")[0], code:file.content})
            }  
        }catch(e){
            alert(e.message, "Error Parsing Gist")
            return
        }
    }catch(e){
        console.log("module is not json")
        // json was not valid, assume we have js
        // check to see if there is a comment that specifies a module name
        //
        let name = null
        if(data.includes("ace.module:")){ 
          try{
              console.log("found ace lable")
            name = JSON.parse(data.split("ace.module:")[1].split("*/")[0]).name
          }catch(e){
              console.log("ace label invalid")
            const url_data=url.split("/")
            name=url_data[url_data.length-1]
          }
        }
        if(!name){
          //either there was no comment to specify a name, or there was an error in reading it
          // we won't overwrite a module unless it is named and the name is something other than "Code"
          // so here, we are going to give it a number to make it unique
            console.log("no name for you")
          let x=1
          while(!!tag(panel_label_to_panel_name("Code "+ x ))){x++}
          files.push({name:"Code " + x, code:data})
        }
    }   

    // now we should have files looking like this
    // files:[{name:module1,content:"function zeta(){..."}, {name:module2,content:"function beta(){..."}]
    // we need to add or update based on the name.
    for(const file of files){
        console.log(file.name,panel_label_to_panel_name(file.name))
        if(!!tag(panel_label_to_panel_name(file.name+" Module"))){
            // a module with this name already exists,  update
            console.log("========= ready to update ============", file.name)
            const editor=ace.edit(panel_label_to_panel_name(file.name) + "_module-content")
            editor.setValue(file.code)
            console.log(editor.getValue())
        }else{
            // no module with this name exists, append
            add_code_module(file.name, file.code)
        }
    }
    
}


function set_css(user_css){
    Jade.css_suffix=user_css
}

function add_library(url){
    // adds a JS library to the head section of the HTML sheet
    const library = document.createElement('script');
    library.setAttribute('src',url);
    console.log("library",library)
    document.head.appendChild(library);
}

function set_theme(theme_name){
    set_style(theme_name)
}

function list_themes(){
    for(const [theme, url] of Object.entries(settings.workbook.styles)){
        console.log(theme, url)
    }
}


function tag(id){
    // a short way to get an element by ID
    return document.getElementById(id)
}

function close_canvas(){
    Jade.panel_stack.pop()
    show_panel(Jade.panel_stack.pop())
}

function open_editor(){
    show_panel(Jade.code_panels[0])
}

function open_examples(){
    show_panel("panel_examples")
}

function open_output(){
    show_panel("panel_output")
}


function open_automations(show_close_button){
    show_automations(show_close_button)
}

function reset(){
    show_panel("panel_home")
}

function show_html(html){
    //A simple function that is mapped differntly for examples than for modules
    //this is the module mapping
    open_canvas("html", html)
}

function open_canvas(panel_name, html, show_panel_close_button, style_name){
    if(style_name){
        set_style(style_name)
    }

    if(!tag(panel_name)){
        build_panel(panel_name)
    }

    if(!Jade.panels.includes(panel_name)){
        Jade.panels.push(panel_name)
    }

    show_panel(panel_name)

    if(html){
        if(show_panel_close_button || show_panel_close_button===undefined){
            tag(panel_name).innerHTML=panel_close_button(panel_name) + html
        }else{
            tag(panel_name).innerHTML= html
        }
        
    }
}
function print(data, heading){
    //if(!header && )
    if(!tag("panel_output").lastChild.lastChild.firstChild.tagName && !heading){
        //no output here, need a headdng
        heading=""
    }
    if(heading!==undefined){
        // there is a header, so make a new block
        console.log("at data")
        const div = document.createElement("div")
        div.className="ace-output"
        const header = document.createElement("div")
        header.className="ace-output-header"  
        const d = new Date()
        let ampm=" am"
        let hours=d.getHours()
        if(hours >11){
            ampm="pm"
            if(hours>12){
                hours=hours-12
            }
        }
        header.innerHTML = '<span class="ace-output-time">' + hours + ":" + ("0"+d.getMinutes()).slice(-2) + ":" + ("0"+d.getSeconds()).slice(-2) + ampm + "</span> " + heading + '<div class="ace-output-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>'
        const body = document.createElement("div")
        body.className="ace-output-body"  
        body.innerHTML = '<div style="margin:0;font-family: monospace;">' + data.replaceAll("\n","<br />")  + "<br />"+ "</div>"
        div.appendChild(header)
        div.appendChild(body)
        tag("panel_output").appendChild(div)
    }else{
        // no header provided, append to most recently added
        tag("panel_output").lastChild.lastChild.firstChild.innerHTML += data.replaceAll("\n","<br />") + "<br />"
    }

  
}
  
function alert(data, heading){
    if(tag("ace-alert")){tag("ace-alert").remove()}
    if(!heading){heading="System Message"}
    const div = document.createElement("div")
    div.className="ace-alert"
    div.id='ace-alert'
    const header = document.createElement("div")
    header.className="ace-alert-header"  
    header.innerHTML = heading + '<div class="ace-alert-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>'
    const body = document.createElement("div")
    body.className="ace-alert-body"  
    body.innerHTML = data
    div.appendChild(header)
    div.appendChild(body)
    document.body.appendChild(div)

}

function show_examples(){
    const panel_name="panel_examples"
    set_style()
    show_panel(panel_name)
  
  }