import { useState } from 'react'
import XLSX from 'xlsx'

function App() {
  const [factory, setFactory] = useState("")
  const [season, setSeason] = useState("")
  const [buymonth, setBuymonth] = useState("")
  const [expData, setExpData] = useState([])
  const [error, setError] = useState([])
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
    if (error.length === 0) {
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
      setExpData([])
    } else {
      let errorStr = ""
      error.forEach(e => {
        errorStr = errorStr + ',' + e
      })
      alert(`工作表名稱:${errorStr}格式錯誤!`)
      setError([])
    }
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
    "2X-LARGE": "Z2XL",
    "3X-LARGE": "Z3XL",
    "4X-LARGE": "Z4XL",
    "5X-LARGE": "Z5XL",
    "6X-LARGE": "Z6XL",
    "2X-LARGE TALL": "Z2XLT",
    "3X-LARGE TALL": "Z3XLT",
    "4X-LARGE TALL": "Z4XLT",
    "5X-LARGE TALL": "Z5XLT",
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
    "40 TALL": "Z40 TALL",
    "42 TALL": "Z42 TALL",
    "44 TALL": "Z44 TALL",
    "46 TALL": "Z46 TALL",
    "48 TALL": "Z48 TALL",
    "50 TALL": "Z50 TALL",
    "52 TALL": "Z52 TALL",
    "54 TALL": "Z54 TALL",
    "56 TALL": "Z56 TALL",
    "58 TALL": "Z58 TALL",
    "60 TALL": "Z60 TALL",
    "40 REGULAR": "Z40 REG",
    "42 REGULAR": "Z42 REG",
    "44 REGULAR": "Z44 REG",
    "46 REGULAR": "Z46 REG",
    "48 REGULAR": "Z48 REG",
    "50 REGULAR": "Z50 REG",
    "52 REGULAR": "Z52 REG",
    "54 REGULAR": "Z54 REG",
    "56 REGULAR": "Z56 REG",
    "58 REGULAR": "Z58 REG",
    "60 REGULAR": "Z60 REG",
    "2 PETITE": "Z2P",
    "4 PETITE": "Z4P",
    "6 PETITE": "Z6P",
    "8 PETITE": "Z8P",
    "10 PETITE": "Z10P",
    "12 PETITE": "Z12P",
    "14 PETITE": "Z14P",
    "16 PETITE": "Z16P",
    "18 PETITE": "Z18P",
    "20 PETITE": "Z20P",
    "2 PETITE SHORT": "Z2PS",
    "4 PETITE SHORT": "Z4PS",
    "6 PETITE SHORT": "Z6PS",
    "8 PETITE SHORT": "Z8PS",
    "10 PETITE SHORT": "Z10PS",
    "12 PETITE SHORT": "Z12PS",
    "14 PETITE SHORT": "Z14PS",
    "16 PETITE SHORT": "Z16PS",
    "18 PETITE SHORT": "Z18PS",
    "20 PETITE SHORT": "Z20PS",
    "16W": "Z16W",
    "18W": "Z18W",
    "20W": "Z20W",
    "22W": "Z22W",
    "24W": "Z24W",
    "26W": "Z26W",
    "28W": "Z28W",
    "30W": "Z30W",
    "28": "Z28",
    "30": "Z30",
    "32": "Z32",
    "34": "Z34",
    "36": "Z36",
    "38": "Z38",
    "40": "Z40",
    "42": "Z42",
    "44": "Z44",
    "46": "Z46",
    "48": "Z48",
    "50": "Z50",
    "52": "Z52",
    "54": "Z54",
    "56": "Z56",
    "58": "Z58",
    "60": "Z60",
    "62": "Z62",
    "64": "Z64",
    "28X30": "Z28X30",
    "30X30": "Z30X30",
    "32X30": "Z32X30",
    "34X30": "Z34X30",
    "36X30": "Z36X30",
    "38X30": "Z38X30",
    "40X30": "Z40X30",
    "42X30": "Z42X30",
    "44X30": "Z44X30",
    "46X30": "Z46X30",
    "48X30": "Z48X30",
    "50X30": "Z50X30",
    "52X30": "Z52X30",
    "54X30": "Z54X30",
    "56X30": "Z56X30",
    "58X30": "Z58X30",
    "60X30": "Z60X30",
    "62X30": "Z62X30",
    "64X30": "Z64X30",
    "28X32": "Z28X32",
    "30X32": "Z30X32",
    "32X32": "Z32X32",
    "34X32": "Z34X32",
    "36X32": "Z36X32",
    "38X32": "Z38X32",
    "40X32": "Z40X32",
    "42X32": "Z42X32",
    "44X32": "Z44X32",
    "46X32": "Z46X32",
    "48X32": "Z48X32",
    "50X32": "Z50X32",
    "52X32": "Z52X32",
    "54X32": "Z54X32",
    "56X32": "Z56X32",
    "58X32": "Z58X32",
    "60X32": "Z60X32",
    "62X32": "Z62X32",
    "64X32": "Z64X32"
  }

  const formatMonthAndDate = (num) => {
    return num < 10 ? '0' + num : num
  }
  const fileChange = (e) => {
    if (factory === '' || season === '' || buymonth === '') {
      alert('請先選擇Factory, Season, Buymonth後再上傳檔案!')
      e.target.value = ''
      return
    }
    const file = e.target.files[0]
    let reader = new FileReader()
    reader.readAsBinaryString(file)
    reader.onload = function (e) {
      let data = e.target.result;
      let wb = XLSX.read(data, { type: 'binary' });
      // let sheet = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])
      // console.log(wb)
      // console.log(sheet)
      // console.log(sheet[11]["__EMPTY"])
      let list = []
      let errorList = []
      setError([])
      for (let name of wb.SheetNames) {
        let sheet = XLSX.utils.sheet_to_json(wb.Sheets[name])
        // console.log(sheet)
        console.log(sheet[1])
        if (sheet[11]) {
          switch (sheet[11]["__EMPTY"]) {
            case "Color Desc":
              // console.log(sheet)
              // Item #
              for (let a = 12; a < sheet.length; a++) {
                let check_0 = Object.hasOwn(sheet[a], " IBO Mass Print")
                if (check_0) {
                  // Color Desc & PPK #
                  for (let b = a; b < sheet.length; b++) {
                    let check_1 = Object.hasOwn(sheet[b], "__EMPTY")
                    let check_2 = Object.hasOwn(sheet[b], "__EMPTY_2")
                    if (check_1 && check_2) {
                      // console.log(`a=${a},b=${b},sheet[${a}][" IBO Mass Print"]:${sheet[a][" IBO Mass Print"]},sheet[${b}]["__EMPTY"]:${sheet[b]["__EMPTY"]},sheet[${b}]["__EMPTY_2"]:${sheet[b]["__EMPTY_2"]}`)
                      let date = new Date()
                      let today = `${date.getFullYear()}.${formatMonthAndDate(date.getMonth() + 1)}.${formatMonthAndDate(date.getDate())}`
                      let shipDate = sheet[3]["__EMPTY"].trim().split("/")
                      let shipDay = `${shipDate[2]}.${shipDate[0]}.${shipDate[1]}`
                      const styleCode = () => {
                        const sizeStr = sheet[b]["__EMPTY_11"]
                        const sizeStrArray = sizeStr.split(" ")
                        const sizeStrLastWord = sizeStr[sizeStr.length - 1]
                        switch (sizeStrLastWord) {
                          case "E":
                            if (sizeStrArray[sizeStrArray.length - 1] === "PETITE") {
                              return "P"
                            } else {
                              return "M"
                            }
                          case "T":
                            if (sizeStrArray[sizeStrArray.length - 1] === "SHORT") {
                              return "P"
                            } else {
                              return "M"
                            }
                          case "L":
                            if (sizeStrArray[sizeStrArray.length - 1] === "TALL") {
                              return "T"
                            } else {
                              return "M"
                            }
                          case "W":
                            return "W"
                          case "X":
                            return "W"
                          default:
                            return "M"
                        }
                      }
                      let row = {
                        "Customer": "JCP",
                        "Customer Name": "",
                        "Order No": `${season.slice(2, 4)}${season.slice(0, 1)}${buymonth.slice(0, 2)}${buymonth.slice(3, 4)}${factory.slice(2, 3)}${sheet[b]["__EMPTY_2"]}${styleCode()}`,
                        "Order Date": today,
                        "Remark": sheet[4]["__EMPTY_8"].trim(),
                        "Season": season,
                        "Division": "KH",
                        "Price Term": "",
                        "Cust.P/O ref.": "",
                        "Cust.P/O Date": "",
                        "Port of Loading": "",
                        "Style": `${sheet[b]["__EMPTY_2"]}${styleCode()}`,
                        "Customer Style": sheet[b]["__EMPTY_2"].toString(),
                        "Description": sheet[b]["__EMPTY_4"],
                        "Qty Unit": "PCS",
                        "Ship Date": shipDay,
                        "Country of Origin": factory,
                        "Ship By": "By Sea",
                        "Ship Description": "USA",
                        "Lot Reference": `${sheet[a][" IBO Mass Print"]}-${sheet[1]["__EMPTY_6"]}`,
                        "Color": sheet[b]["__EMPTY"],
                        "Port of Discharge": "",
                        "Currency": "USD",
                        "BuyMonth": buymonth,
                        "PO Cut": `${sheet[b]["__EMPTY_3"]}`,
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
                      for (let c = b; ; c++) {
                        if (sheet[c]["__EMPTY_11"] && sheet[c]["__EMPTY_24"]) {
                          row[size[sheet[c]["__EMPTY_11"]]] = sheet[c]["__EMPTY_24"]
                        } else if (sheet[c]["__EMPTY_11"] && sheet[c]["__EMPTY_25"]) {
                          row[size[sheet[c]["__EMPTY_11"]]] = sheet[c]["__EMPTY_25"]
                        } else {
                          break
                        }
                      }
                      list.push(row)
                    } else if (sheet[b]["__EMPTY"] === "Item # Subtotal") {
                      break
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
              // console.log(sheet)
              // Pack Item #
              for (let a = 12; a < sheet.length; a++) {
                let check_0 = Object.hasOwn(sheet[a], "__EMPTY")
                if (check_0 && sheet[a]["__EMPTY"] !== "Location Subtotal") {
                  // Color Desc & PPK #
                  for (let b = a; b < sheet.length; b++) {
                    let check_1 = Object.hasOwn(sheet[b], "__EMPTY_9")
                    let check_2 = Object.hasOwn(sheet[b], "__EMPTY_25")
                    if (check_1 && check_2) {
                      // console.log(`a=${a},b=${b},sheet[${a}]["__EMPTY"]:${sheet[a]["__EMPTY"]},sheet[${b}]["__EMPTY_9"]:${sheet[b]["__EMPTY_9"]},sheet[${b}]["__EMPTY_25"]:${sheet[b]["__EMPTY_25"]}`)
                      let date = new Date()
                      let today = `${date.getFullYear()}.${formatMonthAndDate(date.getMonth() + 1)}.${formatMonthAndDate(date.getDate())}`
                      let shipDate = sheet[3]["__EMPTY_1"].trim().split("/")
                      let shipDay = `${shipDate[2]}.${shipDate[0]}.${shipDate[1]}`
                      let styleDesc = sheet[b]["__EMPTY_24"].split(":")
                      let style = styleDesc[0]
                      const styleCode = () => {
                        const sizeStr = sheet[b]["__EMPTY_26"]
                        const sizeStrArray = sizeStr.split(" ")
                        const sizeStrLastWord = sizeStr[sizeStr.length - 1]
                        switch (sizeStrLastWord) {
                          case "E":
                            if (sizeStrArray[sizeStrArray.length - 1] === "PETITE") {
                              return "P"
                            } else {
                              return "M"
                            }
                          case "T":
                            if (sizeStrArray[sizeStrArray.length - 1] === "SHORT") {
                              return "P"
                            } else {
                              return "M"
                            }
                          case "L":
                            if (sizeStrArray[sizeStrArray.length - 1] === "TALL") {
                              return "T"
                            } else {
                              return "M"
                            }
                          case "W":
                            return "W"
                          case "X":
                            return "W"
                          default:
                            return "M"
                        }
                      }
                      let row = {
                        "Customer": "JCP",
                        "Customer Name": "",
                        "Order No": `${season.slice(2, 4)}${season.slice(0, 1)}${buymonth.slice(0, 2)}${buymonth.slice(3, 4)}${factory.slice(2, 3)}${sheet[b]["__EMPTY_9"]}${styleCode()}`,
                        "Order Date": today,
                        "Remark": sheet[4]["__EMPTY_15"].trim(),
                        "Season": season,
                        "Division": "KH",
                        "Price Term": "",
                        "Cust.P/O ref.": "",
                        "Cust.P/O Date": "",
                        "Port of Loading": "",
                        "Style": `${sheet[b]["__EMPTY_9"]}${styleCode()}`,
                        "Customer Style": sheet[b]["__EMPTY_9"].toString(),
                        "Description": style,
                        "Qty Unit": "PCS",
                        "Ship Date": shipDay,
                        "Country of Origin": factory,
                        "Ship By": "By Sea",
                        "Ship Description": "USA",
                        "Lot Reference": `${sheet[b]["__EMPTY"]}-${sheet[1]["__EMPTY_10"].trim()}`,
                        "Color": sheet[b]["__EMPTY_25"],
                        "Port of Discharge": "",
                        "Currency": "USD",
                        "BuyMonth": buymonth,
                        "PO Cut": `${sheet[b]["__EMPTY_8"]}`,
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
                      for (let c = b; ; c++) {
                        if (sheet[c]["__EMPTY_26"]) {
                          row[size[sheet[c]["__EMPTY_26"]]] = sheet[c]["__EMPTY_34"]
                        } else {
                          break
                        }
                      }
                      list.push(row)
                    } else if (sheet[b]["__EMPTY_2"] === "Pack item Subtotal") {
                      break
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
            case "Item #":
              // console.log(sheet)
              // Item #
              for (let a = 12; a < sheet.length; a++) {
                let check_0 = Object.hasOwn(sheet[a], "__EMPTY")
                if (check_0) {
                  // Color Desc & PPK #
                  for (let b = a; b < sheet.length; b++) {
                    let check_1 = Object.hasOwn(sheet[b], "__EMPTY_1")
                    let check_2 = Object.hasOwn(sheet[b], "__EMPTY_5")
                    if (check_1 && check_2) {
                      // console.log(`a=${a},b=${b},sheet[${a}]["__EMPTY"]:${sheet[a]["__EMPTY"]},sheet[${b}]["__EMPTY_1"]:${sheet[b]["__EMPTY_1"]},sheet[${b}]["__EMPTY_5"]:${sheet[b]["__EMPTY_5"]}`)
                      let date = new Date()
                      let today = `${date.getFullYear()}.${formatMonthAndDate(date.getMonth() + 1)}.${formatMonthAndDate(date.getDate())}`
                      let shipDate = sheet[3]["__EMPTY_2"].trim().split("/")
                      let shipDay = `${shipDate[2]}.${shipDate[0]}.${shipDate[1]}`
                      const styleCode = () => {
                        const sizeStr = sheet[b]["__EMPTY_16"]
                        const sizeStrArray = sizeStr.split(" ")
                        const sizeStrLastWord = sizeStr[sizeStr.length - 1]
                        switch (sizeStrLastWord) {
                          case "E":
                            if (sizeStrArray[sizeStrArray.length - 1] === "PETITE") {
                              return "P"
                            } else {
                              return "M"
                            }
                          case "T":
                            if (sizeStrArray[sizeStrArray.length - 1] === "SHORT") {
                              return "P"
                            } else {
                              return "M"
                            }
                          case "L":
                            if (sizeStrArray[sizeStrArray.length - 1] === "TALL") {
                              return "T"
                            } else {
                              return "M"
                            }
                          case "W":
                            return "W"
                          case "X":
                            return "W"
                          default:
                            return "M"
                        }
                      }
                      let row = {
                        "Customer": "JCP",
                        "Customer Name": "",
                        "Order No": `${season.slice(2, 4)}${season.slice(0, 1)}${buymonth.slice(0, 2)}${buymonth.slice(3, 4)}${factory.slice(2, 3)}${sheet[b]["__EMPTY_5"]}${styleCode()}`,
                        "Order Date": today,
                        "Remark": sheet[4]["__EMPTY_19"].trim(),
                        "Season": season,
                        "Division": "KH",
                        "Price Term": "",
                        "Cust.P/O ref.": "",
                        "Cust.P/O Date": "",
                        "Port of Loading": "",
                        "Style": `${sheet[b]["__EMPTY_5"]}${styleCode()}`,
                        "Customer Style": sheet[b]["__EMPTY_5"].toString(),
                        "Description": sheet[b]["__EMPTY_7"],
                        "Qty Unit": "PCS",
                        "Ship Date": shipDay,
                        "Country of Origin": factory,
                        "Ship By": "By Sea",
                        "Ship Description": "USA",
                        "Lot Reference": `${sheet[a]["__EMPTY"]}-${sheet[1]["__EMPTY_14"].trim()}`,
                        "Color": sheet[b]["__EMPTY_1"],
                        "Port of Discharge": "",
                        "Currency": "USD",
                        "BuyMonth": buymonth,
                        "PO Cut": `${sheet[b]["__EMPTY_6"]}`,
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
                      // Size Desc
                      for (let c = b; ; c++) {
                        if (sheet[c]["__EMPTY_16"]) {
                          row[size[sheet[c]["__EMPTY_16"]]] = sheet[c]["__EMPTY_31"]
                        } else {
                          break
                        }
                      }
                      list.push(row)
                    } else if (sheet[b]["__EMPTY_1"] === "Item # Subtotal") {
                      break
                    } else {
                      continue
                    }
                  }
                } else {
                  continue
                }
              }
              break
            default:
              errorList.push(name)
              break
          }
        } else {
          errorList.push(name)
          alert("請移除檔案內空白工作表再進行轉檔!")
        }
      }
      if (errorList.length !== 0) {
        setError([...errorList])
      } else {
        setExpData([...list])
      }
    }
  }
  return (
    <>
      <div>
        <label className="p-4" htmlFor="factory">Factory:</label>
        <select className="border-2 m-2 rounded-md border-lime-500" name="factory" value={factory} onChange={factoryChange}>
          <option value="factory">--select--</option>
          <option value="QVA">QVA</option>
          <option value="QVJ">QVJ</option>
        </select>
        <hr />
        <label className="p-4" htmlFor="season">Season:</label>
        <select className="border-2 m-2 rounded-md border-lime-500" name="season" value={season} onChange={seasonChange}>
          <option value="season">--select--</option>
          <option value="SP24">SP24</option>
          <option value="SU24">SU24</option>
          <option value="FA24">FA24</option>
          <option value="FW24">FW24</option>
          <option value="HO24">HO24</option>
          <option value="SP25">SP25</option>
          <option value="SU25">SU25</option>
          <option value="FA25">FA25</option>
          <option value="FW25">FW25</option>
          <option value="HO25">HO25</option>
          <option value="SP26">SP26</option>
          <option value="SU26">SU26</option>
          <option value="FA26">FA26</option>
          <option value="FW26">FW26</option>
          <option value="HO26">HO26</option>
        </select>
        <hr />
        <label className="p-4" htmlFor="buymonth">BuyMonth:</label>
        <select className="border-2 m-2 rounded-md border-lime-500" name="buymonth" value={buymonth} onChange={buymonthChange}>
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
      <input className="p-4 " type="file" onChange={fileChange} />
      <hr />
      <button className="rounded-md border-2 border-lime-500 p-2 m-4 bg-green-500 text-white" onClick={handleClick}>開始轉檔</button>
    </>
  )
}

export default App