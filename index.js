const path = require('path')
const edge = require('edge-js')

// Compile the C# code and create a JS function
let helloWord = edge.func('helloWord.cs')

// Execute the function (sync)
let result = helloWord({ file: path.join(__dirname, 'test.docx'), text: 'line one\nline two' }, true)

console.log(result)
