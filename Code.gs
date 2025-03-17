function onOpen() {
  let ui = DocumentApp.getUi()
  
  ui.createMenu("Code")
    .addItem('Run code', "runCode")
    .addItem("See last output", "seeLastOutput")
    .addItem("Manage packages", "managePackages")
    .addItem("Manage includes", "manageIncludes")
    .addToUi()
}

function getCode(document) {
    let code = JSON.stringify(document.getBody().getText())

    return code.replaceAll(/[\u201C\u201D]/g, '\\"')
}

function showConsole(html) {
  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Console")
}

function includeDocument(id) {
  return getCode(DocumentApp.openById(id))
}

function runCode() { 
  const convertToArray = str => str.split("\n").map(name => name.trim()).filter(name => name.length)
  
  const packages = JSON.stringify(convertToArray(getArrayValue("packages")))
  let html

  try {
    const includes = JSON.stringify(convertToArray(getArrayValue("includes")).map(id => includeDocument(id)))

    html = `
      <script src="https://cdn.jsdelivr.net/pyodide/v0.23.2/full/pyodide.js"></script>
      <script>
        window.onload = async function() {
          const pyodide = await loadPyodide()

          let result

          try {
            for (const package of ${packages}) {
              await pyodide.loadPackage(package)
            }

            pyodide.runPython("import sys; from io import StringIO; sys.stdout = StringIO();")

            for (const include of ${includes}) {
              pyodide.runPython(eval(include))
            }
            
            pyodide.runPython(${getCode(DocumentApp.getActiveDocument())})
            result = pyodide.runPython('sys.stdout.getvalue()')
          }
          catch(e){
            result = e 
          }

          result = String(result).replaceAll("<", "&lt;").replaceAll("\\n", "<br>")

          google.script.run.storeValue("lastOutput", result)
          document.body.innerHTML = result
        }
      </script>
    `
  }
  catch(e) {
    html = e
  }

  showConsole(html)
}

function storeValue(name, value) {
  PropertiesService.getScriptProperties().setProperty(name, value)
}

function seeLastOutput() { 
  showConsole(PropertiesService.getScriptProperties().getProperty("lastOutput"))
}

function getArrayValue(name) {
  return (PropertiesService.getScriptProperties().getProperty(name) || "")
}

function getListUI(desc, script) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          html, body {
            height: 100%;
            margin: 0;
          }
        </style>
      </head>
      <body style="display: flex; align-items: stretch; justify-content: stretch;">
        <textarea style="width: 100%; height: 100%; box-sizing: border-box;" placeholder="${desc}"></textarea>
        <script>${script}</script>
      </body>
    </html>
  `
}

function managePackages() {
  const html = getListUI("List packages here (line-separated)", `
    google.script.run.withSuccessHandler(packages => {
      document.querySelector("textarea").value = packages
    }).getArrayValue("packages")

    document.querySelector("textarea").addEventListener('input', e => {        
      google.script.run.storeValue("packages", e.target.value)
    })
  `)

  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Packages")
}

function manageIncludes() {
  const html = getListUI("List document IDs here (line-separated)", `
    google.script.run.withSuccessHandler(includes => {
      document.querySelector("textarea").value = includes
    }).getArrayValue("includes")

    document.querySelector("textarea").addEventListener('input', e => {        
      google.script.run.storeValue("includes", e.target.value)
    })
  `)

  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Includes")
}
