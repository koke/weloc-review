#!/usr/bin/env node

if (process.argv.length != 3) {
  console.error("Usage: parse file.xlsx");
  process.exit(1);
}

if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile(process.argv[2]);
var sheet = workbook.Sheets[workbook.SheetNames[0]];
var rows = sheet['!ref'].split(':')[1].match(/\d+/)[0]

var pairs = []
for (var i = 8; i <= rows; i++) {
  var source = sheet['A'+i].v;
  var translation = sheet['B'+i].v;
  pairs.push([source, translation]);
}

function makeIterator(array){
    var nextIndex = 0;

    return {
       next: function(){
         var index = nextIndex++
           return index < array.length ?
               {value: array[index], done: false, index: index} :
               {done: true};
       },
       lastIndex: function() {
         return nextIndex - 1;
       },
       total: array.length
    }
}

var iter = makeIterator(pairs);
function printNextOrDie(iter) {
  var next = iter.next()
  if (next.done) {
    writeAndExit();
  }
  console.log("[%d/%d]\n%s\n%s", next.index + 1, iter.total, next.value[0], next.value[1]);
}

var comments = [];
function writeAndExit() {
  if (comments.length > 0) {
    console.log("Adding %d comments:", comments.length);
    comments.forEach(function(comment) {
      console.log("%d: %s", comment[0], comment[1]);
    });
    XLSX.writeFile(workbook, process.argv[2]);
  }
  process.exit();
}

var readline = require('readline');

var rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  terminal: false
});

printNextOrDie(iter);
rl.on('line', function (cmd) {
  if (cmd != '') {
    comments.push([iter.lastIndex() + 1, cmd]);
    var index = iter.lastIndex() + 8;
    var cell = "D" + index;
    workbook.Sheets[workbook.SheetNames[0]][cell] = {v: cmd};
    console.log("Adding comment to " + cell);
  }
  printNextOrDie(iter);
});

rl.on('close', writeAndExit);
