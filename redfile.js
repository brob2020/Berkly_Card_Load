const reader = require('xlsx')


// Reading our test file
const file = reader.readFile("C:/Users/rober/Desktop/Copy of Pre Paid Card Load ATL Oct 13 300pm.xlsx")
  
let data = []
  
const sheets = file.SheetNames
  
for(let i = 0; i < sheets.length; i++)
{
   const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
      data.push(res)
   })
}
  
// Printing data
console.log(data[4])
 
/* for (let i =0;i<data.length;i++){
    //console.log(i)
    Object.entries(data[i]).forEach(([key, value]) => console.log(`${key}: ${value}`));
    
} */