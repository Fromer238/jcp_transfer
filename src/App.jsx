import { useState } from 'react'
import XLSX from 'xlsx'

function App() {
  const [factory, setFactory] = useState("")
  const [season, setSeason] = useState("")
  const [buymonth, setBuymonth] = useState("")
  const [expData, setExpData] = useState([])
  const factoryChange = (e) => {
    setFactory(e.target.value)
  }
  const seasonChange = (e) => {
    setSeason(e.target.value)
  }
  const buymonthChange = (e) => {
    setBuymonth(e.target.value)
  }
  const originTitleList = ["<Customer(Customer.Code)>", "<Customer/Name>", "<OrderNoPrefix>", "<IssDate(Date)>", "<Style/Shipment/Remark>", "<Season>", "<Division>", "<PrcTerm>", "<CustPORef>", "<CustPODate(Date)>", "<Style/Shipment/PortLoad>", "<Style/Style>", "<Style/CustStyle>", "<Style/Description>", "<Style/Unit>", "<Style/Shipment/ShipDate(Date)>", "<Style/Origin>", "<Style/Shipment/ShipMode>", "<Style/Shipment/ShipDest>", "<Style/Shipment/LotRef>", "<Style/Shipment/Assortment/Color>", "<Style/Shipment/PortDisc>", "<Cur>", "<ExtTerm2>", "<Style/Shipment/ExtDesc1>", "<Style/Shipment/Label>", "<Style/Shipment/ExtDesc6>", "<Style/Shipment/ExtDesc3>", "<Style/ProgramCode>", "<Style/Origin>", "<ExtTerm3>", "<ExtTerm5>", "<Style/Shipment/ExtDesc7>", "<Style/Shipment/ExtDesc8>", "<Style/Shipment/ExtDesc9>"]
  const handleClick = () => {
    // console.log(expData)
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(expData, { origin: "A3" })
    let title = []
    let originTitle = []
    for (let c = 0; ; c++) {
      let cellAddress = XLSX.utils.encode_cell({ r: 2, c: c })
      let cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : ""
      if (cellValue !== "") {
        if (originTitleList[c] !== undefined) {
          originTitle.push(originTitleList[c])
        } else {
          originTitle.push("<Style/Shipment/Assortment/Size,Qty>")
        }

        if (cellValue.slice(0, 1) === "Z") {
          cellValue = cellValue.slice(1)
          title.push(cellValue)
        } else {
          title.push(cellValue)
        }

      } else {
        break
      }
    }
    let row0 = ["Sales Order"]
    XLSX.utils.sheet_add_aoa(worksheet, [row0], { origin: "A1" })
    XLSX.utils.sheet_add_aoa(worksheet, [originTitle], { origin: "A2" })
    XLSX.utils.sheet_add_aoa(worksheet, [title], { origin: "A3" })
    XLSX.utils.book_append_sheet(workbook, worksheet, "sheet1")
    const now = new Date()
    XLSX.writeFileXLSX(workbook, `IG-JCP-${now.getFullYear()}${formatMonthAndDate(now.getMonth() + 1)}${now.getDate()}.xls`)
  }

  const size = {
    "XX-SMALL": "XXS",
    "X-SMALL": "XS",
    "SMALL": "S",
    "MEDIUM": "M",
    "LARGE": "L",
    "X-LARGE": "XL",
    "XX-LARGE": "XXL",
    "XX-SMALL PETITE": "PXXS",
    "X-SMALL PETITE": "PXS",
    "SMALL PETITE": "PS",
    "MEDIUM PETITE": "PM",
    "LARGE PETITE": "PL",
    "X-LARGE PETITE": "PXL",
    "XX-LARGE PETITE": "PXXL",
    "XX-SMALL TALL": "XXST",
    "X-SMALL TALL": "XST",
    "SMALL TALL": "ST",
    "MEDIUM TALL": "MT",
    "LARGE TALL": "LT",
    "X-LARGE TALL": "XLT",
    "XX-LARGE TALL": "XXLT",
    "2": "Z2",
    "4": "Z4",
    "6": "Z6",
    "8": "Z8",
    "10": "Z10",
    "12": "Z12",
    "14": "Z14",
    "16": "Z16",
    "18": "Z18",
    "20": "Z20",
    "0X": "Z0X",
    "1X": "Z1X",
    "2X": "Z2X",
    "3X": "Z3X",
    "4X": "Z4X",
    "5X": "Z5X",
    "2 TALL": "Z2 TALL",
    "4 TALL": "Z4 TALL",
    "6 TALL": "Z6 TALL",
    "8 TALL": "Z8 TALL",
    "10 TALL": "Z10 TALL",
    "12 TALL": "Z12 TALL",
    "14 TALL": "Z14 TALL",
    "16 TALL": "Z16 TALL",
    "18 TALL": "Z18 TALL",
    "20 TALL": "Z20 TALL",
  }

  function formatMonthAndDate(num) {
    return num < 10 ? '0' + num : num
  }
  const fileChange = (e) => {
    const file = e.target.files[0]
    let reader = new FileReader()
    reader.readAsBinaryString(file)
    reader.onload = function (e) {
      let data = e.target.result;
      let wb = XLSX.read(data, { type: 'binary' });
      let sheet = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
      // console.log(sheet)
      let list = []
      switch (sheet[11]["__EMPTY"]) {
        case "Color Desc":
          for (let s = 12; s < sheet.length; s++) {
            let check_reference = Object.hasOwn(sheet[s], " IBO Mass Print")
            if (check_reference) {
              for (let i = s; i < sheet.length; i++) {
                if (sheet[i]["__EMPTY"] === "Item # Subtotal") {
                  break
                }

                let check = Object.hasOwn(sheet[i], "__EMPTY_2")

                if (check) {
                  let date = new Date()
                  let today = `${date.getFullYear()}.${formatMonthAndDate(date.getMonth() + 1)}.${formatMonthAndDate(date.getDate())}`
                  let shipDate = sheet[3]["__EMPTY"].trim().split("/")
                  let shipDay = `${shipDate[2]}.${shipDate[0]}.${shipDate[1]}`
                  let row = {
                    "Customer": "JCP",
                    "Customer Name": "",
                    "Order No": `${season.slice(2, 4)}${season.slice(0, 1)}${buymonth.slice(0, 2)}${buymonth.slice(3, 4)}${factory.slice(2, 3)}${sheet[i]["__EMPTY_2"]}`,
                    "Order Date": today,
                    "Remark": sheet[4]["__EMPTY_8"].trim(),
                    "Season": season,
                    "Division": "KH",
                    "Price Term": "",
                    "Cust.P/O ref.": "",
                    "Cust.P/O Date": "",
                    "Port of Loading": "",
                    "Style": sheet[i]["__EMPTY_2"],
                    "Customer Style": sheet[i]["__EMPTY_2"],
                    "Description": sheet[i]["__EMPTY_4"],
                    "Qty Unit": "PCS",
                    "Ship Date": shipDay,
                    "Country of Origin": factory,
                    "Ship By": "By Sea",
                    "Ship Description": "USA",
                    "Lot Reference": `${sheet[s][" IBO Mass Print"]}-${sheet[1]["__EMPTY_6"]}`,
                    "Color": sheet[i]["__EMPTY"],
                    "Port of Discharge": "",
                    "Currency": "USD",
                    "BuyMonth": buymonth,
                    "PO Cut": `${sheet[i]["__EMPTY_3"]}`,
                    "Label": `${sheet[1]["__EMPTY"].trim()}`,
                    "PD": "",
                    "Assigned Factory": "",
                    "ProgramCode": "",
                    "Factory": factory,
                    "Order Type": "FOB",
                    "Sales Type": "",
                    "PSDD": "",
                    "FPD": "",
                    "LPD": ""
                  }
                  for (let j = i; ; j++) {
                    if (sheet[j]["__EMPTY_11"] && sheet[j]["__EMPTY_24"]) {
                      row[size[sheet[j]["__EMPTY_11"]]] = sheet[j]["__EMPTY_24"]
                    } else if (sheet[j]["__EMPTY_11"] && sheet[j]["__EMPTY_25"]) {
                      row[size[sheet[j]["__EMPTY_11"]]] = sheet[j]["__EMPTY_25"]
                    } else {
                      break
                    }
                  }
                  list.push(row)
                } else {
                  continue
                }
              }

            } else {
              continue
            }
          }
          setExpData([...list])
          break
        case "Pack Item #":
          for (let i = 12; i < sheet.length; i++) {
            let check = Object.hasOwn(sheet[i], "__EMPTY_9")

            if (check) {
              let date = new Date()
              let today = `${date.getFullYear()}.${formatMonthAndDate(date.getMonth() + 1)}.${formatMonthAndDate(date.getDate())}`
              let shipDate = sheet[3]["__EMPTY_1"].trim().split("/")
              let shipDay = `${shipDate[2]}.${shipDate[0]}.${shipDate[1]}`
              let styleDesc = sheet[i]["__EMPTY_24"].split(":")
              let style = styleDesc[0]
              let row = {
                "Customer": "JCP",
                "Customer Name": "",
                "Order No": `${season.slice(2, 4)}${season.slice(0, 1)}${buymonth.slice(0, 2)}${buymonth.slice(3, 4)}${factory.slice(2, 3)}${sheet[i]["__EMPTY_9"]}`,
                "Order Date": today,
                "Remark": sheet[4]["__EMPTY_15"].trim(),
                "Season": season,
                "Division": "KH",
                "Price Term": "",
                "Cust.P/O ref.": "",
                "Cust.P/O Date": "",
                "Port of Loading": "",
                "Style": sheet[i]["__EMPTY_9"],
                "Customer Style": sheet[i]["__EMPTY_9"],
                "Description": style,
                "Qty Unit": "PCS",
                "Ship Date": shipDay,
                "Country of Origin": factory,
                "Ship By": "By Sea",
                "Ship Description": "USA",
                "Lot Reference": `${sheet[i]["__EMPTY"]}-${sheet[1]["__EMPTY_10"].trim()}`,
                "Color": sheet[i]["__EMPTY_25"],
                "Port of Discharge": "",
                "Currency": "USD",
                "BuyMonth": buymonth,
                "PO Cut": `${sheet[i]["__EMPTY_8"]}`,
                "Label": `${sheet[1]["__EMPTY_1"].trim()}`,
                "PD": "",
                "Assigned Factory": "",
                "ProgramCode": "",
                "Factory": factory,
                "Order Type": "FOB",
                "Sales Type": "",
                "PSDD": "",
                "FPD": "",
                "LPD": ""
              }
              for (let j = i; j < sheet.length; j++) {
                if (sheet[j]["__EMPTY_26"]) {
                  row[size[sheet[j]["__EMPTY_26"]]] = sheet[j]["__EMPTY_34"]
                } else {
                  break
                }
              }
              list.push(row)
            } else {
              continue
            }
          }
          setExpData([...list])
          break
        case "Item #":
          for (let a = 12; a < sheet.length; a++) {
            let check_0 = Object.hasOwn(sheet[a], "__EMPTY")
            if (check_0) {
              for (let i = a; i < sheet.length; i++) {
                let check_1 = Object.hasOwn(sheet[i], "__EMPTY_5")
                if (check_1) {
                  let date = new Date()
                  let today = `${date.getFullYear()}.${formatMonthAndDate(date.getMonth() + 1)}.${formatMonthAndDate(date.getDate())}`
                  let shipDate = sheet[3]["__EMPTY_2"].trim().split("/")
                  let shipDay = `${shipDate[2]}.${shipDate[0]}.${shipDate[1]}`
                  let row = {
                    "Customer": "JCP",
                    "Customer Name": "",
                    "Order No": `${season.slice(2, 4)}${season.slice(0, 1)}${buymonth.slice(0, 2)}${buymonth.slice(3, 4)}${factory.slice(2, 3)}${sheet[i]["__EMPTY_5"]}`,
                    "Order Date": today,
                    "Remark": sheet[4]["__EMPTY_19"].trim(),
                    "Season": season,
                    "Division": "KH",
                    "Price Term": "",
                    "Cust.P/O ref.": "",
                    "Cust.P/O Date": "",
                    "Port of Loading": "",
                    "Style": sheet[i]["__EMPTY_5"],
                    "Customer Style": sheet[i]["__EMPTY_5"],
                    "Description": sheet[i]["__EMPTY_7"],
                    "Qty Unit": "PCS",
                    "Ship Date": shipDay,
                    "Country of Origin": factory,
                    "Ship By": "By Sea",
                    "Ship Description": "USA",
                    "Lot Reference": `${sheet[a]["__EMPTY"]}-${sheet[1]["__EMPTY_10"].trim()}`,
                    "Color": sheet[i]["__EMPTY_1"],
                    "Port of Discharge": "",
                    "Currency": "USD",
                    "BuyMonth": buymonth,
                    "PO Cut": `${sheet[i]["__EMPTY_6"]}`,
                    "Label": `${sheet[1]["__EMPTY_2"].trim()}`,
                    "PD": "",
                    "Assigned Factory": "",
                    "ProgramCode": "",
                    "Factory": factory,
                    "Order Type": "FOB",
                    "Sales Type": "",
                    "PSDD": "",
                    "FPD": "",
                    "LPD": ""
                  }
                  for (let j = i; j < sheet.length; j++) {
                    if (sheet[j]["__EMPTY_16"]) {
                      row[size[sheet[j]["__EMPTY_16"]]] = sheet[j]["__EMPTY_31"]
                    } else {
                      break
                    }
                  }
                  list.push(row)
                } else {
                  continue
                }
              }
              setExpData([...list])
            } else {
              continue
            }
          }
          break
      }
    }
  }
  return (
    <>
      <div>
        <label class="p-4" htmlFor="factory">Factory:</label>
        <select class="border-2 m-2 rounded-md border-lime-500" name="factory" value={factory} onChange={factoryChange}>
          <option value="factory">--select--</option>
          <option value="QVA">QVA</option>
          <option value="QVJ">QVJ</option>
        </select>
        <hr />
        <label class="p-4" htmlFor="season">Season:</label>
        <select class="border-2 m-2 rounded-md border-lime-500" name="season" value={season} onChange={seasonChange}>
          <option value="season">--select--</option>
          <option value="SP23">SP23</option>
          <option value="SU23">SU23</option>
          <option value="FW23">FW23</option>
          <option value="SP24">SP24</option>
          <option value="SU24">SU24</option>
          <option value="FW24">FW24</option>
          <option value="SP25">SP25</option>
          <option value="SU25">SU25</option>
          <option value="FW25">FW25</option>
        </select>
        <hr />
        <label class="p-4" htmlFor="buymonth">BuyMonth:</label>
        <select class="border-2 m-2 rounded-md border-lime-500" name="buymonth" value={buymonth} onChange={buymonthChange}>
          <option value="buymonth">--select--</option>
          <option value="01-1">01-1</option>
          <option value="01-2">01-2</option>
          <option value="02-1">02-1</option>
          <option value="02-2">02-2</option>
          <option value="03-1">03-1</option>
          <option value="03-2">03-2</option>
          <option value="04-1">04-1</option>
          <option value="04-2">04-2</option>
          <option value="05-1">05-1</option>
          <option value="05-2">05-2</option>
          <option value="06-1">06-1</option>
          <option value="06-2">06-2</option>
          <option value="07-1">07-1</option>
          <option value="07-2">07-2</option>
          <option value="08-1">08-1</option>
          <option value="08-2">08-2</option>
          <option value="09-1">09-1</option>
          <option value="09-2">09-2</option>
          <option value="10-1">10-1</option>
          <option value="10-2">10-2</option>
          <option value="11-1">11-1</option>
          <option value="11-2">11-2</option>
          <option value="12-1">12-1</option>
          <option value="12-2">12-2</option>
        </select>
      </div>
      <hr />
      <input class="p-4 " type="file" onChange={fileChange} />
      <hr />
      <button class="rounded-md border-2 border-lime-500 p-2 m-4 bg-green-500 text-white" onClick={handleClick}>開始轉檔</button>
    </>
  )
}

export default App
