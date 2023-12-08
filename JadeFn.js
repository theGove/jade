const custom_functions={}
function jadeFn(function_name, args) {
  const fnName="jade_custom_function_"+function_name
  //check to see if we need to reload the custom functions
  
  
  if(!custom_functions[function_name]){    
    // we have no local copy of the function
    const fn=localStorage.getItem(fnName)
    if(fn){
      // the function has been defined, read it in
      custom_functions[function_name]={
        hash:"no hash because we will load it below"
      }
    }else{
      return [[`The JADE custom function named ${function_name} was not found.  You may need to open the JADE add-in again.`]]
    }
  }

  // if we made it here, we have a local copy of the function, check to see if it is up to date
  if(custom_functions[function_name].hash!==localStorage.getItem('hash_' + fnName)){
    // the function is outdated, read it in and update the hash
    const fnText = localStorage.getItem(fnName)
    window[function_name]=eval("(" + fnText + ")")// makes this a window-level function
    custom_functions[function_name].hash=localStorage.getItem('hash_' + fnName)
  }

  // now we know we have the current version of the function

  for(let x=0;x<args.length;x++ ){
    if(args[x].length===1 && args[x][0].length===1){
      args[x]=args[x][0][0]
    }
  }

  let return_value=window[function_name](...args)
  if(!Array.isArray(return_value)){return_value=[[return_value]]}
  return return_value
}

CustomFunctions.associate("fn", jadeFn);
  
  