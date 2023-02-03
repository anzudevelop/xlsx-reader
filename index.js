const xlsx = require('xlsx');
require('dotenv').config()

const LAST_NAME = process.env.LAST_NAME // С маленькой буквы свою фамилию как в компе
var m = new Date().getMonth() + 1
var y = new Date().getYear() % 100
const MONTH = `${m < 10 ? '0' : ''}${m}`//process.env.MONTH
const YEAR = `20${y}`//process.env.YEAR

const notWorkingDaysOfThisYear = process.env.NOT_WORKING_DAYS_OF_THIS_YEAR

const getTimeByPartOfTime = (part) => {
  let h = Math.floor(part * 24)
  let m = Math.floor(((part * 24) % 1) * 60)
  let s = Math.floor(((((part * 24) % 1) * 60) % 1) * 60)
  return {h, m, s}
}

let fixTimeToString = (time) => {
  if (typeof time === 'number') {
      let result = getTimeByPartOfTime(time)
      return `${result.h < 10 ? '0' : ''}${result.h}:${result.m < 10 ? '0' : ''}${result.m}:${result.s < 10 ? '0' : ''}${result.s}`
  }
  return time
}

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

  for(let i = 0; i < parsedData.length; i++)
  {
    if(parsedData[i].hasOwnProperty('Пришел')) parsedData[i]['Пришел'] = fixTimeToString(parsedData[i]['Пришел'])
    if(parsedData[i].hasOwnProperty('Ушел')) parsedData[i]['Ушел'] = fixTimeToString(parsedData[i]['Ушел'])
    if(parsedData[i].hasOwnProperty('Перерыв')) parsedData[i]['Перерыв'] = fixTimeToString(parsedData[i]['Перерыв'])
    if(parsedData[i].hasOwnProperty('Итого')) parsedData[i]['Итого'] = fixTimeToString(parsedData[i]['Итого'])
  }
  // call a function to save the data in a json file
  return(JSON.stringify(parsedData));
}

const dataFromXlsx = JSON.parse(convertExcelFileToJsonUsingXlsx())

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
  let data = dataFromXlsx
  let fullTime = 0
  for(let i = 0; i < data.length - 1; i++) {
    if(data[i].hasOwnProperty('Пришел')) {
      fullTime += executeWorkTimeFromDate(data[i])
    }
  }
  return fullTime
}

const isWorkDay = (d) => {
  let dayWeek = new Date(`${d.slice(3, 5)}.${d.slice(0, 2)}.${d.slice(6, 10)}`).toLocaleString("en-us", { weekday: "short"})
  if(dayWeek != 'Sat' && dayWeek != 'Sun' && !notWorkingDaysOfThisYear.includes(d)) {
    return true
  }
  return false
}

const getWorkDays = () => {
  let data = dataFromXlsx
  let daysCounter = 0
  for(let i = 0; i < data.length - 1; i++) {
    if(data[i].hasOwnProperty('Пришел') && isWorkDay(data[i]['Дата'])) {
      daysCounter++
    }
  }
  return daysCounter
}

const getTotalWorkDaysOfMonth = () => {
  let data = dataFromXlsx
  let daysCounter = 0
  for(let i = 0; i < data.length - 1; i++) {
    if(isWorkDay(data[i]['Дата'])) {
      daysCounter++
    }
  }
  return daysCounter
}

const getTableData = () => {
  let data = dataFromXlsx
  console.log('-> ' + process.env.LAST_NAME + ' <-')
  console.log('---------------------------------------------------------------------------')
  console.log('|     Дата     |    Пришел    |     Ушел     |   Перерыв   |     Итого    |')
  console.log('---------------------------------------------------------------------------')
  for (let i = 0; i < data.length - 1 ; i++) {
    console.log('|  ' + data[i]['Дата'] + '  |   ' + (data[i]['Пришел'] ? data[i]['Пришел'] : '        ') + '   |   ' + (data[i]['Ушел'] ? data[i]['Ушел'] : '        ') + '   |   ' + (data[i]['Пришел'] ? data[i]['Перерыв'] : '        ') + '  |   ' + (data[i].hasOwnProperty('Пришел') ? fixTimeToString(executeWorkTimeFromDate(data[i])) : '00:00:00') + '   |')
  }
  console.log('---------------------------------------------------------------------------')
}

const print = () => {
  //let data = JSON.parse(convertExcelFileToJsonUsingXlsx())
  //let day = 19
  //let dayIndex = findIndexOfDate(data, day)
  
  let partOfMonthTime = getFullMonthTime() - ((8 / 24) * getWorkDays())
  let status = ""
  if (partOfMonthTime > 0) {
    status = "Время переработки:"
  } else {
    status = "Нужно доработать:"
  }
  let timeData = getTimeByPartOfTime(Math.abs(partOfMonthTime))
  let {h, m, s} = timeData

  getTableData()  //Для вывода таблицы со временем

  console.log(status, (partOfMonthTime > 0 ? '\x1b[32m' : '\x1b[31m'), `${(h / 10) >= 1 ? "" : "0"}${h}:${(m / 10) >= 1 ? "" : "0"}${m}:${(s / 10) >= 1 ? "" : "0"}${s}` ,'\x1b[0m')
  console.log(`В этом месяце отработано ${fixTimeToString(getFullMonthTime())} из ${fixTimeToString((8 / 24) * getTotalWorkDaysOfMonth())}`)
}

print()