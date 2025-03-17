function onOpen() {
  let ui = DocumentApp.getUi()
  
  ui.createMenu("Code")
    .addItem('Run code', "runCode")
    .addItem("See last output", "seeLastOutput")
    .addItem("Edit packages", "editPackages")
    .addToUi()
}

function getCode() {
    let code = JSON.stringify(DocumentApp.getActiveDocument().getBody().getText())

    const symbols = {
      "\u201C": '\\"',
      "\u201D": '\\"'
    }

    for (const symbol in symbols) {
      code = code.replaceAll(symbol, symbols[symbol])
    }

    return code
}

function showConsole(html) {
  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Console")
}

function runCode() { 
  const packages = JSON.stringify(getPackages().split("\n").map(name => name.trim()))

  const html = `
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
        
        pyodide.runPython(${getCode()})
        result = pyodide.runPython('sys.stdout.getvalue()')
      }
      catch(e){
        result = e 
      }

      result = String(result).replaceAll("<", "&lt;").replaceAll("\\n", "<br>")

      google.script.run.storeOutput(result)
      document.body.innerHTML = result
    }
  </script>`

  showConsole(html)
}

function storeOutput(result) {
  PropertiesService.getScriptProperties().setProperty("lastOutput", result)
}

function seeLastOutput() { 
  showConsole(PropertiesService.getScriptProperties().getProperty("lastOutput"))
}

function savePackages(packages) {
  PropertiesService.getScriptProperties().setProperty("packages", packages)
}

function getPackages() {
  return (PropertiesService.getScriptProperties().getProperty("packages") || "")
}

function editPackages() {
  const html = HtmlService.createHtmlOutputFromFile("packages.html")

  DocumentApp.getUi().showModalDialog(html, "Packages")
}