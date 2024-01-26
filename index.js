require("dotenv").config();
const reader = require("xlsx");
const puppeteer = require("puppeteer");

// Reading our test file
const file = reader.readFile(
  "C:/Users/rober/Desktop/Copy of Pre Paid Card Load ATL Oct 13 300pm.xlsx"
);
test
let data = [];

const sheets = file.SheetNames;

for (let i = 0; i < sheets.length; i++) {
  const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
  temp.forEach((res) => {
    data.push(res);
  });
}

// Printing data
console.log(data[1].keys());

/*let CaseFileNumber = 0;
let LastName;
let UnitSuite;
let StreetAddress;
let City;
let ProvinceEn;
let PostalCode;
let directive;
let proxy;*/

(async () => {
  for (let i = 0; i < data.length; i++) {
    console.log(i);

    let CaseFileNumber = data[i]["CaseFileNumber"];
    let lastName = data[i]["Primary Beneficiary LastName"];
    let firstName = data[i]["Primary Beneficiary FirstName"];
    let UnitSuite = data[i]["UnitSuite"];
    let StreetAddress = data[i]["StreetAddress"];
    let citie = data[i]["City"];
    let ProvinceEn = data[i]["ProvinceEn"];
    let PostalCode = data[i]["PostalCode"];
    let directive = data[i]["Program Name"];
    let proxy = data[i]["RelatedNumber"];
    console.log(UnitSuite);
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto("https://crc-portal.berkeleypayment.com/#/login", {
      waitUntil: "networkidle0",
    });

    // entering login
    try {
      await page.type(`input[name="Email"]`, process.env.user);
      let p = await page.$('[name="Password"]');
      p.type(process.env.pass);

      await page.waitForTimeout(500);
      //Navigate to next screen
      await Promise.all([
        page.click(`button[type="submit"]`),
        page.waitForNavigation({ waitUntil: "networkidle0" }),
      ]);

      /// Form
      await page.waitForTimeout(500);
      const input = await page.$("#proxy_input", { delay: 100 });
      //console.log(input)
      //page.waitForSelector('#myId');
      input.type(proxy);
      // PROXY
      await page.waitForTimeout(500);
      // First_Name

      let fi = await page.$('input[name="First_Name"]');
      fi.type(firstName);
      await page.waitForTimeout(500);

      // Last_Name
      let li = await page.$('input[name="Last_Name"]');

      await page.waitForTimeout(500);
      li.type(lastName);
      await page.waitForTimeout(500);

     // Insert Program Name 
      const sele = await page.$('input[name="tradeProgram"]');
      sele.click();
      sele.type(directive);
      await page.waitForTimeout(500);

      await page.keyboard.press("Tab");
      //await page.waitForTimeout(5000);

      const adress = await page.$('input[name="Address"]');
      adress.type(StreetAddress);
      await page.waitForTimeout(500);

      const appt = await page.$('input[name="Address2"]');
      if (UnitSuite) {
        console.log("UnitSuite existes");
        appt.type(UnitSuite);
      }

      await page.waitForTimeout(500);

      const city = await page.$('input[name="City"]');
      city.type(citie);
      await page.waitForTimeout(500);
      //page.select('cuntry', 'blue');
      const country = await page.$('[name="cuntry"]');
      country.click();
      country.type("CA");
      await page.waitForTimeout(500);
      const pro = await page.$('[name="provience"]');
      pro.click();
      pro.type("PE");

      await page.waitForTimeout(500);
      const cdp = await page.$('input[name="PostalCode"]');
      cdp.type(PostalCode);
      await page.waitForTimeout(500);

      await page.screenshot({ path: "example.png" });
      await browser.close();
    } catch (error) {
      console.error(error);
      continue;
    }
  }
})();
