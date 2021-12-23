// testing editing a repo and having jsvba pull it from github

function fetch_json(params){
    if(!params.on_fail){params.on_fail=fetch_failed}
    fetch(params.url)
    .then((response) => response.json())
    .then((data) => params.on_success(params,data))
    .catch((error) => params.on_fail(params, error))
}

function html_get_4x(){
    print("at gethtml")
    return `
        <div id="div-4x"  style="border: 1px solid; box-shadow: 5px 5px tan;border-radius:.5rem;max-width:400px; background-color:white;display:none;margin:2rem;padding:1rem">
        <div style="margin-bottom:1rem;font-size:1.5rem;font-weight:bold">Foreign Exchange</div>
        Convert this invoice to currency of the customer's country.<br><br>This requires a free API key from <a href="https://www.alphavantage.co/" target="_blank">Alpha Vantage</a>.<br><br>
            API Key: <input id="api-key-4x" size="10" /> <br><br><button onClick="configure_4x()">Convert</button>
            </div>`
}

function html_toggle_4x(){
    return `
        <div id="div-4x-toggle"  style="border: 1px solid; box-shadow: 5px 5px tan;border-radius:.5rem;max-width:400px; background-color:white;display:none;margin:2rem;padding:1rem">
        <div style="margin-bottom:1rem;font-size:1.5rem;font-weight:bold">Foreign Exchange</div>
        Convert this invoice to currency of the customer's country.<br><br>This invoice is already configured to show it's values in USD or the customer's country currency. <br><br><button onClick="configure_4x()">Toggle Currency</button></div>`
}

function html_invoice(){
    return `
    <div style="border: 1px solid; box-shadow: 5px 5px tan;border-radius:.5rem;max-width:400px; background-color:white;display:inline-block;margin:2rem;padding:1rem">
    <div style="margin-bottom:1rem;font-size:1.5rem;font-weight:bold">Customer Invoice</div>
    Covefront Laboratories conducts DNA analyses for clients all over the world.  Generate a sample invoice for a client.<br><br><button onClick="fill_in_invoice()">Build invoice</button>
    
    </div>`
}

function show_form(){
    let html=[]
    html.push(html_invoice())
    html.push(html_toggle_4x())
    html.push(html_get_4x())

    
  Excel.run(async function(excel){
        const sheet = excel.workbook.worksheets.getActiveWorksheet()
        const inv_rng = sheet.getRange("H1")
        const cur_rng = sheet.getRange("F11:I13")
        inv_rng.load("values")
        cur_rng.load("values")
        await excel.sync()       
        

        open_canvas("forex", html.join(""))

       if(inv_rng.values[0][0]==="INVOICE"){
            if(cur_rng.values[1][0] !== "USD"){
                if(cur_rng.values[1][3]){
                  // already have exchange rate
                    tag("div-4x-toggle").style.display="inline-block"
                }else{
                    tag("div-4x").style.display="inline-block"
                    if(window["api_key_4x"]){
                        tag("api-key-4x").value=window["api_key_4x"] 
                    }
                }
            }
       }

    })
}

function configure_4x(){
       
   if(tag("api-key-4x") && tag("api-key-4x").value){
    window["api_key_4x"]=tag("api-key-4x").value
   }
   
   Excel.run(async function(excel){
        const sheet = excel.workbook.worksheets.getActiveWorksheet()
        const cur_rng = sheet.getRange("F11:I13")
        cur_rng.load("values")
        await excel.sync()

        if(cur_rng.values[1][3]){
            //The worksheet has already been configured for the exchange rate
            if(cur_rng.values[0][0]==="USD"){
                sheet.getRange("F11").values=cur_rng.values[1][0]
                sheet.getRange("H34").numberFormat=cur_rng.values[1][1].replaceAll(".","")+" #,###"
                console.log(cur_rng.values[1][1].replaceAll(".","")+" #,###")
            }else{
                sheet.getRange("F11").values="USD"
                sheet.getRange("H34").numberFormat="$ #,###"
        }
            excel.sync()
            return
            
        }
        
        // the sheet has not yet been configured for forex
        if(cur_rng.values[1][0]==="USD"){
            // The customer is in the US.  No need to configure
            return
        }
        if(! window["api_key_4x"]){
            alert("No API Key Provided")
            return
        }
        fetch_json({
          url:"https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=USD&to_currency="+cur_rng.values[1][0]+"&apikey=" + window["api_key_4x"], 
          on_success:exchange_rate,
          sheet:sheet,
          currency:cur_rng.values[1][0],
          symbol:cur_rng.values[1][1]
        })
    })
    
}


function exchange_rate(params, data){
    console.log("exchange params",params)
    console.log("exchange data",data)
    params.sheet.getRange("I16").copyFrom("G16:G30");
    params.sheet.getRange("F13").copyFrom("F11:H11");
    params.sheet.getRange("I12").values = data["Realtime Currency Exchange Rate"]["5. Exchange Rate"]
    params.sheet.getRange("I13").values = 1
    params.sheet.getRange("F11").values = params.currency
    params.sheet.getRange("G11").formulas="=VLOOKUP(F11,F12:G13,2,FALSE)"
    params.sheet.getRange("H11").formulas="=VLOOKUP(F11,F12:H13,3,FALSE)"
    params.sheet.getRange("G16:G30").formulas="=VLOOKUP($F$11,$F$12:$I$13,4,FALSE)*I16"
    params.sheet.getRange("F12:H13").format.font.color="white"
    params.sheet.getRange('I:I').columnHidden = true
    
    params.sheet.getRange("H34").numberFormat=params.symbol.replaceAll(".","") +" #,###"
    console.log(params.symbol)
    console.log(params.symbol +" #,###")
    params.sheet.context.sync()
    
     tag("div-4x").style.display="none"
     tag("div-4x-toggle").style.display="inline-block"
    
}




function fetch_failed(params, error){
    console.log("failed to get url", params.url)
    console.log("error", error)
}


function fill_in_customer(params, data){
    
    if(tag("div-4x") && data.currency === "USD"){
            tag("div-4x").style.display="none"
    }else{
        tag("div-4x-toggle").style.display="none"
        if(data.state){
            tag("div-4x").style.display="none"
        }else{
            tag("div-4x").style.display="block"
            if(window["api_key_4x"]){
                tag("api-key-4x").value=window["api_key_4x"] 
            }
        }

    }
    
    // include state if present
    const locale=[data.city]
    if(data.state){locale.push(data.state)}
    
    
    locale.push(data.country)
    let d = new Date()
    
    // place customer data
    params.sheet.getRange("F5").values="0"+rand_between(100000,400000) + data.country_code
    params.sheet.getRange("H5").values= d.getFullYear()+"/"+d.getMonth()+"/"+d.getDate()
    params.sheet.getRange("F8").values=data.customer_number
    params.sheet.getRange("H8").values=data.terms
    params.sheet.getRange("A8").values="Attn: " + data.contact
    params.sheet.getRange("A9").values=data.hospital
    params.sheet.getRange("A10").values=locale.join(", ")

    // place currency data
    params.sheet.getRange("G11").values= "$"
    params.sheet.getRange("F11").values= "USD"
    params.sheet.getRange("H11").values= "Dollar"
    

    // hide customer's preferred currency information
    params.sheet.getRange("G12").values= data.symbol
    params.sheet.getRange("F12").values= data.currency
    params.sheet.getRange("H12").values= data.currency_name
    params.sheet.getRange("F12:H13").format.font.color="white"


    // update total to show correct symbol
    //params.sheet.getRange("H34").numberFormat= data.symbol + " 0.00"

    params.sheet.context.sync()


}


function fill_in_product(params, data){
    console.log(params)
    params.sheet.getRangeByIndexes(params.row,0,1,1).values = "["+ data.test_code +"] " + data.test_name
    params.sheet.getRangeByIndexes(params.row,6,1,1).values = data.price
    const qty_index=[1,1,1,1,1,1,1,1,1,1,1,1,2,2,3,4]
    params.sheet.getRangeByIndexes(params.row,5,1,1).values = qty_index[rand_between(0,qty_index.length-1)]


    params.sheet.context.sync()
}



async function fill_in_invoice(sheet_name){

    // get a sheet    
    Excel.run(async function(excel){
        
        if(sheet_name){
            // a worksheet name was provide, use it to get the sheet
            sheet = excel.workbook.worksheets.getItem(sheet_name)
        }else{
            // a worksheet name was not provided, get the active sheet
            sheet = excel.workbook.worksheets.getActiveWorksheet()
        }
        await excel.sync()
        
        // now we have a sheet; check to see if it is a blank invoice.

        const rng_inv = sheet.getRange("H1")
        const rng_cust =  sheet.getRange("A8")
        rng_inv.load("values")
        rng_cust.load("values")
        await excel.sync()
        console.log("rng_inv.values",rng_inv.values)
        console.log("rng_cust.values",rng_cust.values)
        console.log("array",rng_cust.values[0][0])
        if(rng_inv.values[0][0]==="INVOICE" && rng_cust.values[0][0]===""){
            // it appears to be a blank invoice.  build
            console.log("ready to complete the invoice")
            fetch_json({
              url:"https://thegove.github.io/data/hospital/"+rand_between(1,200)+".json", 
              on_success:fill_in_customer,
              sheet:sheet
            })
            
            const prod_count=rand_between(1,9)
            const products=[]
            for(let i=1;i<=prod_count;i++){
                const prod=rand_between(1,127)
                if(!products.includes(prod)){
                    products.push(prod)
                }
            }    
            console.log(products)
            
            for(let i=0;i<products.length;i++){
                fetch_json({
                  url: "https://thegove.github.io/data/dna_test/"+products[i]+".json", 
                  on_success: fill_in_product,
                  sheet: sheet,
                  row: i+15
                })
            }
            
            
            
        }else{
            // this is not a blank invoice, build an invoice and try again
            console.log("building")
            sheet_name = await Excel.run(build_empty_invoice_example)
            console.log((sheet_name))
            fill_in_invoice(sheet_name)
        }
    })


}


async function build_empty_invoice_example(excel){
  /*ace.listing:{"name":"Build Invoice","description":"Makes a blank invoice."}*/
  
  const worksheets=excel.workbook.worksheets
  worksheets.load("items")
  await excel.sync();
  
  for(const ws of worksheets.items){ws.load("name")}
  await excel.sync();
  
  let file_number=1
  let proposed_name="Invoice"
  
  do{
   let name_found = false
   for(const ws of worksheets.items){
      if(ws.name.toLowerCase()===proposed_name.toLowerCase()){
          name_found = true
          break
      }
    }
    if(!name_found){break}
    proposed_name = "Invoice " + file_number++
  }while(file_number < 200)


  if(file_number >= 200){
    return // we have too many invoices, let's not go crazy
  }

  const sheet = excel.workbook.worksheets.add(proposed_name)
  sheet.activate()
  await excel.sync()


  // add data
  
  // company information
  sheet.getRange("A1:A4").values=[
      ["Covefront Laboratories"],
      ["698 Candlewood Lane"],
      ["Cabot Cove, Maine 04456"],
      ["Phone: (207) 555-1212"]
  ]
  
  // invoice labels
  sheet.getRange("H1").values="INVOICE"
  sheet.getRange("A7").values="Bill To"
  sheet.getRange("A15").values="DESCRIPTION"
  sheet.getRange("F4").values="INVOICE #"
  sheet.getRange("H4").values="DATE"
  sheet.getRange("F7").values="CUSTOMER ID"
  sheet.getRange("H7").values="TERMS"
  sheet.getRange("F10").values="Symbol"
  sheet.getRange("H10").values="Currency"
  sheet.getRange("A15").values="DESCRIPTION"
  sheet.getRange("A31").values="Thank you for your business!"
  sheet.getRange("A37").values="If you have any questions about this invoice, please contact"
  sheet.getRange("A38").values="Martin Eddington  (207) 555-1212  meddington@covefrontanalytic.com"
  sheet.getRange("F31:F34").values=[
      ["SUBTOTAL"],
      ["TAX RATE"],
      ["TAX"],
      ["TOTAL"]
      ]
  sheet.getRange("F15:H15").values=[
      ["QTY","PRICE","AMOUNT"]
      ]

  // set column widths    
  sheet.getRange("A1").format.columnWidth=40
  sheet.getRange("B1").format.columnWidth=75
  sheet.getRange("C1").format.columnWidth=115
  sheet.getRange("D1").format.columnWidth=40
  sheet.getRange("E1").format.columnWidth=80
  sheet.getRange("F1").format.columnWidth=35
  sheet.getRange("G1").format.columnWidth=45
  sheet.getRange("H1").format.columnWidth=90
  sheet.getRange("I1").format.columnWidth=70

  // set the height of row 1
  sheet.getRange("a1").getEntireRow().format.rowHeight=40
  
  // set font size
  sheet.getRange("A1").format.font.size=20
  sheet.getRange("H1").format.font.size=35
  sheet.getRange("F34:H34").format.font.size=16


  // set cell fill color
  sheet.getRange("A1:J50").format.fill.color="white"
  sheet.getRange("F31:F34").format.fill.color="lightSteelBlue"
  sheet.getRange("H31:H34").format.fill.color="#DDDDDD"
  
  // set cell font color
  sheet.getRange("A1").format.font.color="mediumBlue"
  sheet.getRange("H1").format.font.color="maroon"

  // set cell indentation
  sheet.getRange("A2:A4").format.indentLevel=1
  sheet.getRange("A8:A13").format.indentLevel=1
  sheet.getRange("A7").format.indentLevel=2
  sheet.getRange("A15").format.indentLevel=2
  sheet.getRange("F31:F34").format.indentLevel=2
  
  // set several properties of a cell with a function
  make_header(sheet, "A7")
  make_header(sheet, "A15")
  make_header(sheet, "F15:H15")
  make_header(sheet, "F4:H4")
  make_header(sheet, "F7:H7")
  make_header(sheet, "F10:H10")
  
  // merge cells
  sheet.getRange("A7:C7").merge()
  sheet.getRange("A15:E15").merge()
  sheet.getRange("F31:G34").merge(true)
  sheet.getRange("F4:G10").merge(true)
  sheet.getRange("A31:E31").merge()
  sheet.getRange("A37:H38").merge(true)
  
  // center cells
  sheet.getRange("F4:H15").format.horizontalAlignment = "Center"
  sheet.getRange("F16:F30").format.horizontalAlignment = "Center"
  sheet.getRange("A31:A38").format.horizontalAlignment = "Center"
  
  // right align
  sheet.getRange("H1").format.horizontalAlignment = "Right"
  
  // set cells bold
  sheet.getRange("A1:H1").format.font.bold = true
  sheet.getRange("F34:H34").format.font.bold = true
  
  // borders
  let borders = sheet.getRange("A30:H30").format.borders
  borders.getItem("EdgeBottom").style="Continuous"

  borders = sheet.getRange("A16:H30").format.borders
  borders.getItem("InsideHorizontal").color="lightGrey"
  borders.getItem("InsideHorizontal").style="dot"
  borders.getItem("InsideHorizontal").weight="hairline"
  
  borders = sheet.getRange("E16:H30").format.borders
  borders.getItem("InsideVertical").color="lightGrey"
  borders.getItem("InsideVertical").style="dot"
  borders.getItem("InsideVertical").weight="hairline"
  
  // add formulas
  sheet.getRange("H16:H30").formulas = '=if(G16="","",product(F16:G16))'
  sheet.getRange("H31").formulas = "=sum(H16:H30)"
  sheet.getRange("H32").values = 0.06
  sheet.getRange("H33").formulas = "=product(H31:H32)"
  sheet.getRange("H34").formulas = "=H31+H33"
  
  // set cell number formats
  sheet.getRange("G16:H33").numberFormat="#,###"
  sheet.getRange("H34").numberFormat="$ #,###"
  sheet.getRange("H32").numberFormat="0.0#%"

  await excel.sync();
  return proposed_name
}

function make_header(sheet, range){
     sheet.getRange(range).format.fill.color="royalBlue"
     sheet.getRange(range).format.font.color="white"
     sheet.getRange(range).format.font.bold=true
     sheet.getRange(range).format.font.size=12
}
function setup(){
    show_html('<button onclick="Excel.run(build_empty_invoice_example)">Build full empty invoice</button>')
}

function rand_between(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}


function launch_fill_in_invoice(){
/*ace.listing:{"name":"Enter Invoice Data","description":"Fills the invoice with sample data."}*/
    fill_in_invoice()
}

function launch_show_form(){
      /*ace.listing:{"name":"Show invoice form","description":"Displays a form to manage the invoice example."}*/
  show_form()
}