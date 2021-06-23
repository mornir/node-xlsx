const ExcelJS = require("exceljs")
const sanityClient = require("@sanity/client")
const toPlainText = require("./toPlainText")

const workbook = new ExcelJS.Workbook()

workbook.creator = 'VKF'

const fileName = 'Terminofeu_Export.xlsx'

const worksheet = workbook.addWorksheet("Deutsch")

worksheet.columns = [
  { header: "Begriff", key: "term", width: 32 },
  { header: "Definition", key: "definition", width: 62 },
  { header: "Anmerkung", key: "note", width: 62 },
]

const client = sanityClient({
  projectId: "nipfx4rq",
  dataset: "production",
  apiVersion: "2021-06-19", // use current UTC date - see "specifying API version"!
  useCdn: false, // `false` if you want to ensure fresh data
})

const query = /* groq */ `*[_type == 'entry' && status == "definition"] {deTitle, 'definition': content.de.definition, 'note': content.de.note}`

client
  .fetch(query)
  .then((entries) => {
    entries.forEach((entry) => {
      const note = entry.note ? toPlainText(entry.note) : ''
      const definition = entry.definition ? toPlainText(entry.definition) : ''
      worksheet.addRow({
        term: entry.deTitle,
        definition,
        note,
      })
    })

    return workbook.xlsx.writeFile(fileName)
  })
  .then(() => {
    console.log("success")
  })
  .catch((e) => {
    console.log(e)
  })
