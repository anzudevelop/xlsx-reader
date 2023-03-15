const fs = require("fs");

const readValue = fs.readFileSync('date.json', 'utf8')
const year = Number(readValue.slice(1, 5))
const month = Number(readValue.slice(6, 8))
const day = Number(readValue.slice(9, 11))
const h = Number(readValue.slice(12,14))
const m = Number(readValue.slice(15,17))
const s = Number(readValue.slice(18,20))

function formatDate(date) {
    return date.getFullYear() + '/' +
      (date.getMonth()<10?`0${date.getMonth() + 1}`:(date.getMonth() + 1)) + '/' +
      (date.getDate()<10?`0${date.getDate()}`:date.getDate()) + ' ' +
      (date.getHours()<10?`0${date.getHours()}`:date.getHours()) + ':' +
      (date.getMinutes()<10?`0${date.getMinutes()}`:date.getMinutes()) + ':' +
      (date.getSeconds()<10?`0${date.getSeconds()}`:date.getSeconds());
   }

let date = new Date()

if((date.getMonth() + 1) > month || date.getDate() > day) fs.writeFileSync("date.json", JSON.stringify(formatDate(date)));
else {
    if(date.getHours() < h || (date.getHours() == h && date.getMinutes < m)) fs.writeFileSync("date.json", JSON.stringify(formatDate(date)));
}
console.log(formatDate(date))