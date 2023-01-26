const xlsx = require('xlsx');
var fs = require('fs');
const path = require('path')
require('dotenv').config()

const LAST_NAME = process.env.LAST_NAME // С маленькой буквы свою фамилию как в компе
var m = new Date().getMonth() + 1
var y = new Date().getYear() % 100
const MONTH = `${m < 10 ? '0' : ''}${m}`//process.env.MONTH
const YEAR = `20${y}`//process.env.YEAR

function convertExcelFileToJsonUsingXlsx() {

  // Read the file using pathname
  const file = xlsx.readFile(`//srv3/WorkTimes/${LAST_NAME}/${LAST_NAME}_${MONTH}-${YEAR}.xlsx`);

  // Grab the sheet info from the file
  const sheetNames = file.SheetNames;
  const totalSheets = sheetNames.length;

  // Variable to store our data
  let parsedData = [];

  // Loop through sheets
  for (let i = 0; i < totalSheets; i++) {

      // Convert to json using xlsx
      const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[i]]);

      // Skip header row which is the colum names
      tempData.shift();

      // Add the sheet's json to our data array
      parsedData.push(...tempData);
  }

 // call a function to save the data in a json file

 return(JSON.stringify(parsedData));
}

const getTimeByPartOfTime = (part) => {
  let h = Math.floor(part * 24)
  let m = Math.floor(((part * 24) % 1) * 60)
  let s = Math.floor(((((part * 24) % 1) * 60) % 1) * 60)
  return {h, m, s}
}

const findIndexOfDate = (data, day) => {
  for(let i = 0; i < data.length; i++) {
    if(data[i]["Дата"].split(".")[0] == day) return i
  }
}

const executeWorkTimeFromDate = (day) => {
  let secondsInDay = 24 * 60 * 60
  let h, m, s
  h = Number(day['Пришел'].split(':')[0]) * 3600
  m = Number(day['Пришел'].split(':')[1]) * 60
  s = Number(day['Пришел'].split(':')[2])
  let startWork = h + m + s
  h = Number(day['Перерыв'].split(':')[0]) * 3600
  m = Number(day['Перерыв'].split(':')[1]) * 60
  s = Number(day['Перерыв'].split(':')[2])
  let lunchTime = h + m + s
  h = Number(day['Ушел'].split(':')[0]) * 3600
  m = Number(day['Ушел'].split(':')[1]) * 60
  s = Number(day['Ушел'].split(':')[2])
  let endDay = h + m + s
  return ((endDay - lunchTime - startWork) / secondsInDay)
}

const getFullMonthTime = () => {
  let data = JSON.parse(convertExcelFileToJsonUsingXlsx())
  let fullTime = 0
  for(let i = 0; i < data.length - 1; i++) {
    if(data[i]["Итого"] != 0 || data[i].hasOwnProperty('Пришел')) {
      fullTime += data[i]["Итого"]
      if(data[i]["Итого"] == 0) {
        fullTime += executeWorkTimeFromDate(data[i])
      }
    }
  }
  return fullTime
}

const isWorkDay = (d) => {
  let dayWeek = new Date(`${d.slice(3, 5)}.${d.slice(0, 2)}.${d.slice(6, 10)}`).toLocaleString("en-us", { weekday: "short"})
  if(dayWeek != 'Sat' && dayWeek != 'Sun') {
    return true
  }
  return false
}

const getWorkDays = () => {
  let data = JSON.parse(convertExcelFileToJsonUsingXlsx())
  let daysCounter = 0
  for(let i = 0; i < data.length - 1; i++) {
    if((data[i]["Итого"] != 0 || data[i].hasOwnProperty('Пришел')) && isWorkDay(data[i]['Дата'])) {
      daysCounter++
    }
  }
  return daysCounter
}

const print = () => {
  //let data = JSON.parse(convertExcelFileToJsonUsingXlsx())
  //let day = 19
  //let dayIndex = findIndexOfDate(data, day)
  
  let partOfMonthTime = getFullMonthTime() - ((8 / 24) * getWorkDays())
  let status = ""
  if (partOfMonthTime > 0) {
    status = "Время переработки: "
  } else {
    status = "Нужно доработать: "
    partOfMonthTime *= -1
  }
  let timeData = getTimeByPartOfTime(partOfMonthTime)
  let {h, m, s} = timeData
  console.log(`${status}${(h / 10) > 1 ? "" : "0"}${h}:${(m / 10) > 1 ? "" : "0"}${m}:${(s / 10) > 1 ? "" : "0"}${s}`)
}

print()