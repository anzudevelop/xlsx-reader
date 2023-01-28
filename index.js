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

  let parsedData = [{"Дата":"02.01.2023","Итого":"00:00:00"},{"Дата":"03.01.2023","Итого":"00:00:00"},{"Дата":"04.01.2023","Итого":"00:00:00"},{"Дата":"05.01.2023","Итого":"00:00:00"},{"Дата":"06.01.2023","Итого":"00:00:00"},{"Дата":"07.01.2023","Итого":"00:00:00"},{"Дата":"08.01.2023","Итого":"00:00:00"},{"Дата":"09.01.2023","Пришел":"08:35:15","Ушел":"12:01:29","Перерыв":"00:00:00","Итого":"03:26:14","Примечания":"Дата прихода взята по данным с сервера"},{"Дата":"10.01.2023","Пришел":"07:48:30","Ушел":"16:52:03","Перерыв":"00:30:00","Итого":"08:33:32","Примечания":""},{"Дата":"11.01.2023","Пришел":"09:37:03","Ушел":"18:10:14","Перерыв":"00:30:00","Итого":"08:03:10","Примечания":"Дата прихода взята по данным с сервера"},{"Дата":"12.01.2023","Пришел":"07:37:18","Ушел":"16:07:53","Перерыв":"00:30:00","Итого":"08:00:35","Примечания":"Дата прихода взята по данным с сервера"},{"Дата":"13.01.2023","Пришел":"07:15:09","Ушел":"11:19:20","Перерыв":"00:30:00","Итого":"03:34:11","Примечания":""},{"Дата":"14.01.2023","Пришел":"08:01:24","Ушел":"11:56:01","Перерыв":"00:00:00","Итого":"03:54:37","Примечания":""},{"Дата":"15.01.2023","Итого":"00:00:00"},{"Дата":"16.01.2023","Пришел":"08:27:29","Ушел":"10:30:46","Перерыв":"00:00:00","Итого":"02:03:16","Примечания":""},{"Дата":"17.01.2023","Пришел":"07:30:35","Ушел":"16:02:53","Перерыв":"00:30:00","Итого":"08:02:18","Примечания":""},{"Дата":"18.01.2023","Пришел":"07:05:45","Ушел":"10:43:00","Перерыв":"00:00:00","Итого":"03:37:15","Примечания":""},{"Дата":"19.01.2023","Пришел":"10:14:13","Ушел":"18:00:00","Перерыв":"00:30:00","Итого":"07:15:47","Примечания":"Дата ухода взята по последней записи логов. Сообщите точное время ухода ответственому лицу"},{"Дата":"20.01.2023","Пришел":"07:56:56","Ушел":"16:31:35","Перерыв":"00:30:00","Итого":"00:00:00","Примечания":"Дата прихода взята по данным с сервера"},{"Дата":"21.01.2023","Пришел":"07:02:44","Ушел":"11:04:32","Перерыв":"00:30:00","Итого":"00:00:00","Примечания":"Дата ухода взята по последней записи логов"},{"Дата":"22.01.2023","Итого":"00:00:00"},{"Дата":"23.01.2023","Пришел":"08:15:02","Ушел":"17:30:21","Перерыв":"00:30:00","Итого":"00:00:00","Примечания":"Дата прихода взята по данным с сервера"},{"Дата":"24.01.2023","Пришел":"08:02:12","Ушел":"17:45:46","Перерыв":"00:30:00","Итого":"00:00:00","Примечания":""},{"Дата":"25.01.2023","Пришел":"08:00:23","Ушел":"17:37:29","Перерыв":"00:30:00","Итого":"00:00:00","Примечания":""},{"Дата":"26.01.2023","Пришел":"08:17:38","Ушел":"17:30:54","Перерыв":"00:30:00","Итого":"00:00:00","Примечания":""},{"Дата":"27.01.2023","Итого":"00:00:00"},{"Дата":"28.01.2023","Итого":"00:00:00"},{"Дата":"29.01.2023","Итого":"00:00:00"},{"Дата":"30.01.2023","Итого":"00:00:00"},{"Дата":"31.01.2023","Итого":"00:00:00"},{"Итого":"56:30:57"}]

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
    console.log('|  ' + data[i]['Дата'] + '  |   ' + (data[i]['Пришел'] ? data[i]['Пришел'] : '        ') + '   |   ' + (data[i]['Ушел'] ? data[i]['Ушел'] : '        ') + '   |   ' + (data[i]['Пришел'] ? data[i]['Перерыв'] : '        ') + '  |   ' + (data[i]['Итого'] == '00:00:00' ? data[i].hasOwnProperty('Пришел') ? fixTimeToString(executeWorkTimeFromDate(data[i])) : '00:00:00' : '00:00:00' ) + '   |')
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
    status = "Время переработки: "
  } else {
    status = "Нужно доработать: "
    partOfMonthTime *= -1
  }
  let timeData = getTimeByPartOfTime(partOfMonthTime)
  let {h, m, s} = timeData

  getTableData()  //Для вывода таблицы со временем

  let response = ''
  response += `${status}${(h / 10) > 1 ? "" : "0"}${h}:${(m / 10) > 1 ? "" : "0"}${m}:${(s / 10) > 1 ? "" : "0"}${s}\n`
  response += `В этом месяце отработано ${fixTimeToString(getFullMonthTime())} из ${fixTimeToString((8 / 24) * getTotalWorkDaysOfMonth())}\n`

  console.log(response)
}

print()