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
  function tag(id){
    // a short way to get an element by ID
    return document.getElementById(id)
  }
  