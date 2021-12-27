
class Jade{

  // Class Properties
  static settings={}
  static css_suffix=""
  static panels=['panel_home','panel_examples']
  static panel_labels=["Home", "Examples", "Output"]
  static code_panels=[]
  static panel_stack=['panel_home']


  // Class Methods exposed to be called by Jade Users


  static async load_gist(gist_id){
  // We can use gist as repository for the code then either
  // consume it or import it.  Gists have multiple files
  // that become modules in jsvba

    console.log("gisting", gist_id)
    try{
        
      const response = await fetch(`https://api.github.com/gists/${gist_id}?${new Date()}`)
      const data = await response.json()
      for(const file of Object.values(data.files)){
          console.log("===================================")
          console.log(file.content)
          console.log("===================================")
          Jade.incorporate_code(file.content)
      }
      auto_exec()
      Jade.incorporate_code("auto_exec=null")
      
    }catch(e){
      console.log("Error fetching gist", e)
    }
  }
  static async import_code_module(url_or_gist_id){
      Jade.settings.workbook.module_to_import=url_or_gist_id
      console.log("at import code mod", Jade.settings.workbook)

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
            while(!!tag(Jade.panel_label_to_panel_name("Code "+ x ))){x++}
            files.push({name:"Code " + x, code:data})
          }
      }   

      // now we should have files looking like this
      // files:[{name:module1,content:"function zeta(){..."}, {name:module2,content:"function beta(){..."}]
      // we need to add or update based on the name.
      for(const file of files){
          console.log(file.name,Jade.panel_label_to_panel_name(file.name))
          if(!!tag(Jade.panel_label_to_panel_name(file.name+" Module"))){
              // a module with this name already exists,  update
              console.log("========= ready to update ============", file.name)
              const editor=ace.edit(Jade.panel_label_to_panel_name(file.name) + "_module-content")
              editor.setValue(file.code)
              console.log(editor.getValue())
          }else{
              // no module with this name exists, append
              Jade.add_code_module(file.name, file.code)
          }
      }
      
  }
  static set_css(user_css){
      Jade.css_suffix=user_css
  }
  static add_library(url){
      // adds a JS library to the head section of the HTML sheet
      const library = document.createElement('script');
      library.setAttribute('src',url);
      console.log("library",library)
      document.head.appendChild(library);
  }
  static close_canvas(){
      Jade.panel_stack.pop()
      Jade.show_panel(Jade.panel_stack.pop())
  }
  static open_editor(){
      Jade.show_panel(Jade.code_panels[0])
  }
  static open_output(){
      Jade.show_panel("panel_output")
  }
  static open_automations(show_close_button){
      Jade.show_automations(show_close_button)
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
      if(style_name){
          Jade.set_style(style_name)
      }

      if(!tag(panel_name)){
          Jade.build_panel(panel_name)
      }

      if(!Jade.panels.includes(panel_name)){
          Jade.panels.push(panel_name)
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
      if(heading!==undefined){
          // there is a header, so make a new block
          console.log("at data")
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
          body.innerHTML = '<div style="margin:0;font-family: monospace;">' + data.replaceAll("\n","<br />")  + "<br />"+ "</div>"
          div.appendChild(header)
          div.appendChild(body)
          tag("panel_output").appendChild(div)
      }else{
          // no header provided, append to most recently added
          tag("panel_output").lastChild.lastChild.firstChild.innerHTML += data.replaceAll("\n","<br />") + "<br />"
      }

    
  }
  static open_examples(){
      const panel_name="panel_examples"
      Jade.set_style()
      Jade.show_panel(panel_name)
  }

  // Class methods that still need work before shown to the public

  static set_theme(theme_name){
    Jade.set_style(theme_name)
  }
  static list_themes(){
      for(const [theme, url] of Object.entries(Jade.settings.workbook.styles)){
          console.log(theme, url)
      }
  }


  // Class Methods NOT meant to be called by Jade Users
  // It's not really a problem if they do, we just don't
  // think they are useful and we don't document them.


  static officeReady(info){// invoked when the office addin infrastructure has loaded
      if (info.host === Office.HostType.Excel) {
          Excel.run(async (excel)=>{
            const xl_settings = excel.workbook.settings.getItemOrNullObject("jade").load("value");
            //  const code_module_ids_from_settings = excel.workbook.settings.getItemOrNullObject("code_module_ids").load("value");
            await excel.sync()
            if(xl_settings.isNullObject){
              // no Jade.settings so let's configuire some defaults
              Jade.settings.user={
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
                  fontSize:"14pt",
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
                  useSoftTabs:true,
                  navigateWithinSoftTabs:false,
                  tabSize:2,
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
              Jade.settings.workbook={
                code_module_ids:[],
                examples_gist_id:"904983747625c3fdc8dfa69e0aaa0f08",
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
              //console.log("xl_settings",xl_settings.value)
              Jade.settings.workbook=xl_settings.value.workbook
              Jade.settings.user = xl_settings.value.user
            }// if Jade.settings null object
            console.log("before start_me_up, Jade.settings", Jade.settings)
            Jade.configure_settings()
            Jade.start_me_up()
          })
      }else{
          document.getElementById("sideload-msg").style.display = "flex"
          document.getElementById("menu").style.display = "none"
      }
  }
  static start_me_up(){
  //console.log("at start_me_up")
  Jade.settings.workbook.styles.system=tag("head_style").innerText
  Jade.panels.push("panel_home")
  
  //load code from one gist if specified.  
  console.log("about ot load")
  if(Jade.settings.workbook.load_gist_id){
    console.log("in if")
    load_gist(Jade.settings.workbook.load_gist_id)
    
  }
  
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
    for(const panel_name of Jade.code_panels){
      tag(panel_name + "_editor-page").style.height = Jade.editor_height()
    }
  }, true);
  Jade.init_examples()
  Jade.init_output()
  
  // ---------------- Initializing Code Editors -----------------------------
  if(Jade.settings.workbook.code_module_ids.length>0){// show the button to view code modules
    Jade.show_element("open-editor")
  }
  //console.log("at init_code_editors       Jade.settings.workbook.code_module_ids", Jade.settings.workbook.code_module_ids)
  Excel.run(async (excel)=>{
    const parser = new DOMParser();
    for(const code_module_id of Jade.settings.workbook.code_module_ids){
      const xmlpart=excel.workbook.customXmlParts.getItem(code_module_id)
      const xmlBlob = xmlpart.getXml();
      await excel.sync()
      
      const doc = parser.parseFromString(xmlBlob.value, "application/xml");
      const module_name=doc.getElementsByTagName("name")[0].textContent
      const module_code = atob(doc.getElementsByTagName("code")[0].textContent)
      //const settings=atob(doc.getElementsByTagName("settings")[0].textContent)// might want to rename
      const options=atob(doc.getElementsByTagName("options")[0].textContent)
      console.log("just loaded module", module_name)
      //console.log("Jade.settings2", JSON.parse(Jade.settings))
      //console.log("options", options)
      //console.log("options-parsed", JSON.parse(options))
      Jade.add_code_editor(module_name, module_code,code_module_id, null, JSON.parse(options))        
    }
  })


  console.log("end of  start_me_up, Jade.settings", Jade.settings)
  }
  static configure_settings(){
    console.log("Jade.settings", Jade.settings)
    //if(!tag('settings-page').className.includes("hidden")){
      //tag('jade-theme').focus();
      
      //console.log("fontSize", Jade.settings.user.ace_options.fontSize)  
      tag("jade-font-size").value = Jade.settings.user.ace_options.fontSize.replace("pt","")
      if(Jade.settings.user.ace_options.wrap===false){
        tag("jade-word-wrap").value="no-wrap"
      }else if(Jade.settings.user.ace_options.indentedSoftWrap){
        tag("jade-word-wrap").value="wrap-indented"
      }else{
        tag("jade-word-wrap").value="wrap"
      }
      //console.log("theme", Jade.settings.user.ace_options.theme)
      //console.log("theme", Jade.settings.user.ace_options.theme.split("/")[2])
      tag("examples-gist-id").value  = Jade.settings.workbook.examples_gist_id
      tag("jade-theme").value  = Jade.settings.user.ace_options.theme.split("/")[2]
      tag("jade-line-numbers").checked = Jade.settings.user.ace_options.showGutter
      if(Jade.settings.workbook.load_gist_id){
        tag("load-gist-id").value = Jade.settings.workbook.load_gist_id
      }
    //}
  }
  static async save_settings(){
    if(tag("examples-gist-id").value && Jade.settings.workbook.examples_gist_id!==tag("examples-gist-id").value){
      // different examples specified.  Need to rebuild
      Jade.rebuild_examples(tag("examples-gist-id").value)
    }
    Jade.settings.workbook.examples_gist_id=tag("examples-gist-id").value
    Jade.settings.workbook.load_gist_id=tag("load-gist-id").value
    Jade.settings.user.ace_options.theme="ace/theme/" + tag("jade-theme").value
    Jade.settings.user.ace_options.fontSize=tag("jade-font-size").value + "pt"
    Jade.settings.user.ace_options.showGutter=tag("jade-line-numbers").checked
    switch(tag("jade-word-wrap").value){
      case "wrap":
        Jade.settings.user.ace_options.wrap=true
        Jade.settings.user.ace_options.indentedSoftWrap=false
        break
      case "wrap-indented":
        Jade.settings.user.ace_options.wrap=true
        Jade.settings.user.ace_options.indentedSoftWrap=true
        break
      default:  
        Jade.settings.user.ace_options.wrap="off"
    }
    //console.log(Jade.settings.user.ace_options)
    Jade.apply_editor_options(Jade.settings.user.ace_options)
  
    await Jade.write_settings_to_workbook()
  
    Jade.hide_element('settings-page')
  }
  static async write_settings_to_workbook(){
    console.log("at Jade.write_settings_to_workbook", Jade.settings)
    await Excel.run(async (excel)=>{
      const xl_settings = excel.workbook.settings;
      xl_settings.add("jade", Jade.settings);  // adds or sets the value
      await excel.sync()
    })
  }
  static apply_editor_options(options){
    for(const panel_name of Jade.code_panels){
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
    for(const panel_name of Jade.code_panels){
      console.log("panel_name",panel_name,Jade.panel_label_to_panel_name(name))
      if(panel_name === Jade.panel_label_to_panel_name(name) + "_module"){
        //we have a match, and that's a no-no
        alert('A module named "'+name+'" already exists in this workbook.  <br><br>Choose a differnt name.',"Invalid Module Name")
        return
      }
    }


    if(!code){// no code is pased in, determine which default code to import
      if(Jade.code_panels.length === 0){
        code = Jade.default_code()
      }else{  
        code = Jade.default_code("panel_" + name.toLowerCase().split(" ").join("_") + "_module")
      }
    }



    Jade.add_code_editor(name, code, "")
    Jade.hide_element("add-module")
    Jade.show_element("open-editor")
    Jade.show_panel(Jade.code_panels[Jade.code_panels.length-1])
    Jade.write_module_to_workbook(code,Jade.code_panels[Jade.code_panels.length-1])
  }
  static async get_style(style_name, url, integrate_now){

    if(integrate_now===undefined){
       integrate_now=true
    }
  
    if(!Jade.settings.workbook.styles[style_name]){
      Jade.settings.workbook.styles[style_name]=url
    }
  
    if(Jade.settings.workbook.styles[style_name].substr(0,8)==="https://"){
      // the style has not yet been fetched
      //console.log("fetching")
      const response = await fetch(Jade.settings.workbook.styles[style_name])
      const data = await response.text()
      //console.log("data",data)
      Jade.settings.workbook.styles[style_name]=data
      if(integrate_now){
        document.getElementById("head_style").remove()
        document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + Jade.settings.workbook.styles[style_name] + Jade.css_suffix + "</style>")
      }
    }
    
  }
  static set_style(style_name){

    //console.log("at set style", style_name)
    let css_sfx = Jade.css_suffix
    if(!style_name){
      style_name="system"
      css_sfx=""
  
    }
  
    if(Jade.settings.workbook.styles[style_name].substr(0,8)==="https://"){
      // this style has not been fetched.  Get it now
      //console.log("in iff")
      Jade.get_style(style_name)
    }
  
    const style_tag = document.getElementById("head_style")
    if(style_tag.dataset.name!==style_name ){
      // only update the style tag if it is a differnt name
      style_tag.remove()
      document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + Jade.settings.workbook.styles[style_name] + css_sfx + "</style>")
    }
    
  }
  static panel_close_button(panel_name){
    return '<div id="close_' + panel_name + '" onclick="Jade.close_canvas(\'' + panel_name + '\')" class="top-corner" style="padding:5px 5px 0 5px;margin:5px 15px 0 0; cursor:pointer"><i class="far fa-window-close fa-2x"></i></i></div>'
  }
  static build_panel(panel_name, show_close_button){
    const div=document.createElement("div");
    div.id=panel_name
    div.style.display="none"
    div.innerHTML=Jade.panel_close_button(panel_name)
    if(!Jade.panels.includes(panel_name)){
      Jade.panels.push(panel_name)
    }
    document.body.appendChild(div)
    if (show_close_button===undefined){
      show_close_button=true
    }
    if(!show_close_button){
      Jade.hide_element("close_" + panel_name)
    }
  }
  static show_automations(show_close_button){
    const panel_name="panel_listings"
  
    // get the list of functions
    //###################################################### need to iterate over all modules
    const html=['<h2 style="margin:0 0 0 1rem">Active Automations</h2><ol>']
    //console.log("code panels", Jade.code_panels)
    for(const code_panel of Jade.code_panels){
      const editor = ace.edit(code_panel + "-content");
      const code = editor.getValue();
      const parsed_code=Jade.parse_code(code)
  
  
     //console.log(parsed_code)
      
  
      for(const element of parsed_code.body){
        let call_stmt = null
        if(element.type==="FunctionDeclaration"){
          if(element.id && element.id.name){
            // this is a named function
            if(element.params.length===0){
              //there are no params. it is callable
              call_stmt=element.id.name + "()"
            } else if(element.params.length===1){
              // there is one param.  
              if(element.async){
                if(("excel ctx context").includes(element.params[0].name)){
                  // this is an async function with a single parameter named excel, ctx or context.  Run by passing context
                  call_stmt = "Excel.run(" + element.id.name + ")"
                }
              }
            }  
            if(call_stmt){ // this is a function we can run directly
              // check for comment
              const function_text = window[ element.id.name]+''
             //console.log(function_text)
  
  
              if(function_text.includes("Jade.listing:")){ // this is a function we can run directly and it as the comment
                //console.log("found a comment", func)
                const comment = function_text.split("Jade.listing:")[1].split("*/")[0]
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
  
    if(html.length===1){
      //Did not find any properly configured functions
      html.push("There are currently no active automations in this workbook.")
    }else{
      html.push("</ul>")  
    }
    Jade.open_canvas("panel_listings",html.join(""), show_close_button)
  }
  static show_panel(panel_name){

    if(Jade.code_panels.includes(panel_name)){
      // set the size in case it is off
      if(tag(panel_name + "_function-names").length===0){
        // there are no function to run
        Jade.hide_element(panel_name + "_function-names")
      }
      tag(panel_name + "_editor-page").style.height = Jade.editor_height()
      try{
        ace.edit(panel_name + "-content").focus()
      }catch(e){
        ;console.log("could not access ace.  This is expected",e)
      }
    }
    //################## 3 is  a problsm
    if(Jade.panels.slice(0, 3).includes(panel_name) || Jade.code_panels.includes(panel_name)){
      Jade.set_style()
    }
    
   //console.log("trying",panel_name)
    for(const panel of Jade.panels){
      if(panel===panel_name){
       //console.log("showing", panel)
        if(tag("selector_"+ panel_name)){
          tag("selector_"+ panel_name).value=panel_name
        }
        tag(panel).style.display="block"  
        Jade.panel_stack.push(panel)
      }else{
        //console.log(" hiding", panel)
        tag(panel).style.display="none"  
      }
    }
  
    if(Jade.code_panels.includes(panel_name)){
      //focus the ace editor
      try{
        ace.edit(panel_name + "-content").focus()
      }catch(e){
        ;console.log("could not access ace.  This is expected",e)
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
    // Jade.settings are things gove is storing with the module
    // options are the options from the ace editor
    console.log("Jade.settings", Jade.settings)
    let options = Jade.settings.user.ace_options
  
  console.log(1)
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
    Jade.code_panels.push(panel_name)
    Jade.panel_labels.push(Jade.panel_name_to_panel_label(panel_name))
    Jade.panels.push(panel_name)
    Jade.build_panel(panel_name, false)
    tag(panel_name).dataset.module_name = module_name
    tag(panel_name).dataset.module_xmlid = module_xmlid
    
    console.log(2)
  
   //console.log("initializing examples", tag(panel_name))
    
    tag(panel_name).appendChild(Jade.get_panel_selector(panel_name))
    const editor_container = document.createElement("div")
    editor_container.className=panel_name+"-content"
  
    let div = document.createElement("div");
    div.style.padding=".2rem"
    div.style.verticalAlign="middle"
    div.innerHTML = '<button title="Save code to workbook" onclick="Jade.update_editor_script(\'' + panel_name + '\')">Save</button> <button  title="Save code to workbook and execute" onclick="Jade.code_runner(tag(\'' + panel_name + '_function-names' + '\').value,\'' + panel_name + '\')">Run</button> <select id="' + panel_name + '_function-names"></select>';
    div.style.height="22px"
    div.style.fontFamily="auto";
    div.style.fontSize = "1rem";
    div.style.padding=".2rem"
    div.style.backgroundColor = "#eee";
    div.id=panel_name + "_editor-bar"
    editor_container.appendChild(div);
    console.log(3)
  
  
    //console.log("=======================================")
    //console.log(div);
    //console.log(div.clientHeight);
  
    const box = document.createElement("div");
    box.id = panel_name + "_editor-page";
    box.style.width = "100%";
    box.style.height = Jade.editor_height()
    box.style.display = "inline-block";
    box.style.position = "relative";
    console.log(4)
  
   //console.log("document",document.body.clientHeight);
   //console.log("scr",tag("panel_code_editor").Height);
  
  
    div = document.createElement("div");
    div.id = panel_name + "-content";
    div.dataset.edited=false
    div.innerHTML = code.toHtmlEntities();
  
    box.appendChild(div);
    console.log(5)
  
  
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
    console.log(6)
  
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
  
      //console.log("Jade.settings", Jade.settings)
      editor.moveCursorTo(mod_settings.cursorPosition.row, mod_settings.cursorPosition.column)
  
      editor.commands.addCommand({  // toggle word wrap
        name: "wrap",
        bindKey: {win: "Alt-z", mac: "Alt-z"},
        exec: function(editor) {
          for(const panel of Jade.code_panels){
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
          for(const panel of Jade.code_panels){
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
          for(const panel of Jade.code_panels){
            if(tag(panel).style.display==="block"){
              // we found the one that is visible
              Jade.code_runner(tag(panel + '_function-names').value, panel)
              break//exit the loop
            }
          }
        }
      })
  
    });
  
    
  
    editor_container.style.display = "block";
  
    const parsed_code=Jade.parse_code(code)
    if(!parsed_code.error){
      // only incorporate the code if it is free of syntax errors
      Jade.incorporate_code(code)  
    }
  
    
    
    tag(panel_name).appendChild(editor_container)
  
    Jade.load_function_names_select(code, panel_name)
    if(mod_settings.func){
      tag(panel_name + "_function-names").value=mod_settings.func
  
    }
  
    // AutoExecutable function
   //console.log("about to autoexec")
    try{
     //console.log("in try")
      auto_exec()
    }catch(e){
     ;console.log("catch",e)
    }  
  }
  static code_runner(script_name,panel_name){
    //console.log(script_name,panel_name)
    if (tag(panel_name + "-content").dataset.edited==="true"){
      if(!Jade.update_editor_script(panel_name)){
        // Jade.update_editor_script returns false if there is a 
        // syntax error.  Don't run the old code
        return
      }
    }
    
    if(script_name.includes("(excel)")){
      setTimeout("Excel.run(" + script_name.split("(")[0] + ")", 0) //run the function
    }else{
      setTimeout(script_name, 0) //run the function
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
    Jade.fill_examples()
  }  
  static fill_examples(gist_ids){
    //gist_ids is a comma delimited list of gist ids that hold examples
    
    if(!gist_ids){gist_ids=Jade.settings.workbook.examples_gist_id}
    const panel_name="panel_examples"
    const panel=tag(panel_name)
   //console.log("initializing examples")
    panel.appendChild(Jade.get_panel_selector(panel_name))
    const div = document.createElement("div")
    div.className="content"
    div.id="e_content"
    panel.appendChild(div)
    console.log("gist_ids",gist_ids, Jade.settings)
  
    const gist_list=[]
    console.log("gist_ids.split(",")",gist_ids.split(","))
    for(const gist of gist_ids.split(",")){
      console.log("gisting",gist)
      gist_list.push(gist.trim())
    }
    console.log("gist_list",gist_list)
    const html=[]
    for(const gist of gist_list){
      html.push(`<div id="${gist}"></div>`)
    }
    tag("e_content").innerHTML = html.join("");
  
    for(let i=0;i<gist_list.length;i++){
      console.log("gist_list[i]",gist_list[i])
      Jade.get_example_html(gist_list[i],i+1)// get the gist and integrate the examples
    }
  }
  static get_example_html(gist_id, sequence){

    const url=`https://api.github.com/gists/${gist_id}?${Date.now()}`
    const gist_url=`https://gist.github.com/${gist_id}`
    console.log("building examples",url)
    
    console.log("about to fetch",url)
    fetch(url)
    .then((response) => response.text())
    .then((json_text) => {
    //  console.log("json_text",json_text)
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
          // there is a block comment.  assume it is a descriptions
          html.push(temp.split("*/")[0].split("/*")[1])  
        }
        html.push(`</p><div id="page${sequence*100+example_number}"></div>`)
      }
      tag(gist_id).innerHTML = html.join("");
  
      return
  
    });
  
  
  }
  static show_example(id, url, lines) {
    // place the code in an editable box for user to see and play with
    // these examples should be made in script lab to have the right format

    const elem = tag("page" + id);
    //console.log(id + id);
    if (elem.innerHTML === "") {
      console.log("lines", lines)
      elem.innerHTML='<img id="loading-image" width="50" src="assets/loading.gif" />'
  
      fetch(url)
        .then((response) => response.text())
        .then((data) => {
         //console.log(data)
          //const gist = jsyaml.load(data);
          //console.log(gist)
          tag("loading-image").remove()
          let div = document.createElement("div");
          div.id="example" + id + "_html"
          //div.innerHTML = gist.template.content;
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
  
          elem.style.display = "block";
          Jade.incorporate_code(Jade.show_example_html_script(id))
  
          Jade.incorporate_code(data)
          setup() // setup must be defined in the example
          if(!Jade.is_visible(tag("page_" + id ))){
            tag("page_" + id ).scrollIntoView(false)
          }
        })
        .catch((error) => {
         ;console.log(error);
        })
    } else {
      if (elem.style.display === "block") {
        elem.style.display = "none";
      } else {
        elem.style.display = "block";
        Jade.incorporate_code(Jade.show_example_html_script(id))
        Jade.incorporate_code(ace.edit("editor" + id).getValue())
        setup() // setup must be defined in the example
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
    const parsed_code=Jade.parse_code(code)
  
    if(parsed_code.error){
      alert(parsed_code.error, "Syntax Error")
      editor.moveCursorTo(parsed_code.error.split("(")[1].split(":")[0]-1,parsed_code.error.split("(")[1].split(":")[1].split(")")[0])
      editor.focus()
      return false
    }
  
  
    Jade.incorporate_code(Jade.show_example_html_script(id))
    Jade.incorporate_code(code)
    
    return true
  }    
  static parse_code(code){
    try{
      return acorn.parse(code, { "ecmaVersion": 8 }  );
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
    console.log("about to write module")
    Jade.write_module_to_workbook(code, panel_name)
  
    // update the height of the editor in case it has gotten out of synch
    tag(panel_name + "_editor-page").style.height = Jade.editor_height()
  
  
    // show_html is defined differntly for modules than for examples
    // we need to be sure it is defined correctly for modules, so we set it here
    Jade.incorporate_code('function show_html(html){open_canvas("html", html)}')
  
  
    //Check for syntax errors
    const parsed_code=Jade.parse_code(code)
  
    if(parsed_code.error){
      alert(parsed_code.error, "Syntax Error")
      editor.moveCursorTo(parsed_code.error.split("(")[1].split(":")[0]-1,parsed_code.error.split("(")[1].split(":")[1].split(")")[0])
      editor.focus()
  
      return false
    }
    
    // load the user's code into the browser
    Jade.incorporate_code(code)
  
    // put the function names in the function dropdown
    Jade.load_function_names_select(parsed_code, panel_name)
    
    return true
  }
  static incorporate_code(code){
    // It turns out that the script block does not need to persist in the HTML
    // once the script block is loaded, the JS is parsed and not again referenced.
    // So, we create a script block, append it to the document body, then remove.
    const script = document.createElement("script");
    script.innerHTML = code
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
      Jade.settings.cursorPosition = editor.getCursorPosition()
      options=editor.getOptions()
    }catch(e){
      //console.log("This is an expected error: ace editor not yet built",e)
    }
  
    //console.log("writing options", JSON.stringify(options))
    
    const module_name = tag(panel_name).dataset.module_name
    let xmlid = tag(panel_name).dataset.module_xmlid
  
    if (xmlid){  // workbook as already been saved and has and xmlid
      //console.log("saving an existing book")
      Jade.save_module_to_workbook(code, module_name, Jade.settings, xmlid, options)
    }else{  // workbook not yet saved, function call will return an xmlid
      Jade.save_module_to_workbook(code, module_name, Jade.settings, xmlid, options, tag(panel_name))
    }
  }
  static save_module_to_workbook(code, module_name, mod_settings, xmlid, options, tag_to_hold_new_xml_id){
    // options are currently being ignored, they are handled globally instead of at the module level
    // save to workbook
    Excel.run(async (excel)=>{
      //console.log("saving code", panel_name) 
      if(!mod_settings){
        mod_settings = {
          cursorPosition:{row: 0, column: 0},
        }
      }
      //The next line has been disabled because we are not currently maintaining options at the module level
      //const module_xml = "<module xmlns='http://schemas.gove.net/code/1.0'><name>"+module_name+"</name><Jade.settings>"+btoa(JSON.stringify(Jade.settings))+"</Jade.settings><options>"+btoa(JSON.stringify(options))+"</options><code>"+btoa(code)+"</code></module>"
  
      // save module without options
      const module_xml = "<module xmlns='http://schemas.gove.net/code/1.0'><name>"+module_name+"</name><Jade.settings>"+btoa(JSON.stringify(mod_settings))+"</Jade.settings><options>"+btoa(null)+"</options><code>"+btoa(code)+"</code></module>"
      if(xmlid){
        console.log("updating xml", xmlid, typeof xmlid)
        const customXmlPart = excel.workbook.customXmlParts.getItem(xmlid);
        customXmlPart.setXml(module_xml)
        excel.sync()
        console.log("------- launched saving: existing -------")
        
      }else{
        //console.log("creating xml")
        const customXmlPart = excel.workbook.customXmlParts.add(module_xml);
        customXmlPart.load("id");
        await excel.sync();
  
        //console.log("customXmlPart",customXmlPart.getXml())
        // this is a newly created module and needs to have a custom xmlid part made for it
        console.log("23443", Jade.settings, customXmlPart.id)
        Jade.settings.workbook.code_module_ids.push(customXmlPart.id)                   // add the id to the list of ids
        Jade.write_settings_to_workbook()
        console.log("------- launched saving: newly created -------")
        console.log(typeof tag_to_hold_new_xml_id, tag_to_hold_new_xml_id)
        if(tag_to_hold_new_xml_id){
          tag_to_hold_new_xml_id.dataset.module_xmlid = xmlid
        }
      }
  
    })
  
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
            option_value=element.id.name + "()"
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
    if(!Jade.panel_labels.includes(panel_label)){
      Jade.panel_labels.push(panel_label)
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
    for (let i=0; i<Jade.panel_labels.length; i++) {
      var option = document.createElement("option");
      option.value = Jade.panel_label_to_panel_name(Jade.panel_labels[i]) 
     //console.log("-->", option.value)
      option.text = Jade.panel_labels[i];
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
  static select_page(){
    console.log("panel name",window.event.target.value)
    Jade.show_panel(window.event.target.value)
  }
  static show_element(tag_id){
    // removes the hidden class from a tag's css
    if (typeof tag_id==="string"){
      var the_tag=tag(tag_id)
    }else{  
      the_tag=tag_id
    }
    the_tag.className=the_tag.className.replaceAll("hidden","")
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
    console.log("tag_id", tag_id)
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
    console.log("at show import", Jade.settings.workbook.module_to_import)
    if(Jade.settings.workbook.module_to_import){
      console.log("in if")
      tag("gist-url").value=Jade.settings.workbook.module_to_import
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
}// end of Jade class


/*** Convert a string to HTML entities */
String.prototype.toHtmlEntities = function() {
  return this.replace(/./gm, function(s) {
      // return "&#" + s.charCodeAt(0) + ";";
      return (s.match(/[a-z0-9\s]+/i)) ? s : "&#" + s.charCodeAt(0) + ";";
  });
};


String.prototype.toTitleCase=function() {
  str = this.toLowerCase().split(' ');
  for (var i = 0; i < str.length; i++) {
    str[i] = str[i].charAt(0).toUpperCase() + str[i].slice(1); 
  }
  return str.join(' ');
}

function tag(id){
  // a short way to get an element by ID
  return document.getElementById(id)
}
function alert(data, heading){
  if(tag("jade-alert")){tag("jade-alert").remove()}
  if(!heading){heading="System Message"}
  const div = document.createElement("div")
  div.className="jade-alert"
  div.id='jade-alert'
  const header = document.createElement("div")
  header.className="jade-alert-header"  
  header.innerHTML = heading + '<div class="jade-alert-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>'
  const body = document.createElement("div")
  body.className="jade-alert-body"  
  body.innerHTML = data
  div.appendChild(header)
  div.appendChild(body)
  document.body.appendChild(div)

}
