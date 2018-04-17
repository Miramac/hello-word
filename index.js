const path = require('path')
const edge = require('edge-js')

// Compile C# and create JS Function
let helloWord = edge.func('helloWord.cs')

// Execute the function (sync)
let result = helloWord({ file: path.join(__dirname, 'test.docx'), text: 'line one\nline two' }, true)

console.log(result)
