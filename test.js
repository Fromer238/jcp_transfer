function formatMonth(month) {
    return month < 10 ? '0' + month : month;
}
  
const currentDate = new Date();
const currentMonth = formatMonth(currentDate.getMonth() + 1);
  
console.log(currentMonth);