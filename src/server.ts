import puppeteer, {Browser} from "puppeteer";
import { createWorker } from 'tesseract.js';
import fs from 'fs/promises'
import * as XLSX from 'xlsx';

// const url = "https://www.businessmap.burgas.bg/bg/business?item=3672";
const baseUrl = "https://www.businessmap.burgas.bg/bg/business?item=";

type CompanyData = {
  name: string
  website?: string
  contactPersonName: string
  emailImage?: string
  address: string // Fixed typo: "adress" → "address"
  phones?: string[]
  employeesNumber?: number
  mainActivity: string
  hasEmployeesData: boolean
  additionalInfo?: {
    /**
     * @description на кои континенти работи
     */
    businessTerritory: string,
    /**
     * @description в кои държави работи
     */
    countries: string[]
  }[],
  email?: string
}

type CompanyDataExcel = {
  name: string
  website?: string
  contactPersonName: string
  address: string
  phones?: string
  employeesNumber?: number
  mainActivity: string
  hasEmployeesData: boolean,
  BusinessTerritory?: string,
  countries?: string,
  email?: string
}

const main = async () => {
  const browser = await puppeteer.launch();
  let companies: CompanyData[] = []
  const worker = await createWorker('eng')
  for (let i = 0; i < 3673; i++) {
    const url = `${baseUrl}${i}`
    console.log('url :', url);
    companies.push(await scrapePage(url, browser))
  }
  companies = await Promise.all(
    companies.map(async (company) => {
      if(!company.emailImage) return company

      const email: string =
        (await worker.recognize(company.emailImage)).data.text.replace(/\n/g, '').replace(/\s+/g, '').trim() || ""
      company.email = email
      return company
    }),
  )
  fs.writeFile(`./public/scrape-${Date.now()}.json`, JSON.stringify(companies), 'utf8')
  saveToExcel(companies)
  await worker.terminate()
  await browser.close();
  // const companiesWithErrors = 
};


async function scrapePage(url: string, browser: Browser) {
  const page = await browser.newPage()

  await page.goto(url, { waitUntil: "domcontentloaded" })

  const companyData: CompanyData = await page.evaluate(() => {
    // Extract data in browser context
    const name = document.querySelector("h1")?.textContent?.trim() || ""
    const website =
      document
        .querySelector('div.average-color > a[target="_blank"]')
        ?.getAttribute("href") || ""
    const contactPersonName = document.querySelector("div.average-color > header:nth-child(1)")
      ?.textContent?.trim() || ""
    const mainActivity = document
        .querySelector("article.animate:nth-child(2) > p:nth-child(5)")
        ?.textContent?.trim() || ""
    const address = document.querySelector("h2")?.textContent?.trim() || ""
    const phones = Array.from(document.querySelectorAll('a[href^="tel:"]')).map(phoneElement => phoneElement?.textContent?.trim() || "")
    const emailImage = document.querySelector(".email-image")?.getAttribute('src') || ""
    const hasEmployeesData =
      document
        .querySelector(".block > dl:nth-child(1) > dt:nth-child(1)")
        ?.innerHTML?.trim() === "Брой служители"
    let employeesNumber = null
    if(hasEmployeesData) employeesNumber = document.querySelector("div.block.animate > dl > dd")?.innerHTML?.trim();

    const additionalInfo: CompanyData["additionalInfo"] = Array.from(
      document.querySelectorAll("ul.bullets"),
    ).map((section) => {
      const businessTerritory =
        section.querySelector("h3")?.innerText?.trim() || ""
      const countries: CompanyData["additionalInfo"][0]["countries"] =
        Array.from(section.querySelectorAll("li")).map((activity) =>
          activity?.innerText?.trim(),
        )
      return {
        businessTerritory,
        countries,
      }
    })

    const companyData: CompanyData = {
      name,
      website,
      address,
      mainActivity,
      contactPersonName,
      phones,
      employeesNumber,
      hasEmployeesData: hasEmployeesData,
      additionalInfo,
      emailImage
    }
    // Return data to Node.js context
    return companyData
  })
  
  page.close()
  return companyData
}
function saveToExcel(data: CompanyData[]) {
  const workSheetData = data.map(company => {
    const countriesFlat = company.additionalInfo.flatMap(info => info.countries).join('; ') || ""
    const businessTerritoryFlat = company.additionalInfo.flatMap(info => info.businessTerritory).join("; ") || ""
    
    const excel: CompanyDataExcel = {
      name: company.name,
      website: company.website || "",
      contactPersonName: company.contactPersonName,
      address: company.address,
      phones: company.phones.join(", "),
      employeesNumber: company.employeesNumber,
      mainActivity: company.mainActivity,
      BusinessTerritory: businessTerritoryFlat,
      countries: countriesFlat,
      hasEmployeesData: company.hasEmployeesData,
      email: company.email,
    }
    return excel
  })
  console.dir(workSheetData, { depth: null })
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(workSheetData)
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Companies');
  XLSX.writeFile(workbook, `./public/companies-${Date.now()}.xlsx`)

}
main();