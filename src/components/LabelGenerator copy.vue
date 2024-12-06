<template>
  <div class="p-6">
    <h1 class="text-2xl font-bold mb-4">Générateur d'Étiquettes</h1>
    <div class="mb-4">
      <label class="block text-gray-700">Télécharger le fichier Excel :</label>
      <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" class="mt-2" />
    </div>
    <div class="mb-4">
      <label class="block text-gray-700">Entrez le numéro :</label>
      <input type="number" v-model.number="userNumber" class="mt-2 p-2 border rounded w-full" />
    </div>
    <button
      @click="generateLabels"
      class="bg-blue-500 text-white px-4 py-2 rounded disabled:opacity-50"
      :disabled="isGenerating"
    >
      Générer les Étiquettes
    </button>
    <div v-if="errorMessage" class="mt-4 text-red-500">
      {{ errorMessage }}
    </div>
    <div v-if="isGenerating" class="mt-4">
      Génération des étiquettes en cours, veuillez patienter...
    </div>
    <div v-if="downloadLink" class="mt-4">
      <a
        :href="downloadLink"
        download="etiquettes.zip"
        class="bg-green-500 text-white px-4 py-2 rounded"
      >
        Télécharger les Étiquettes
      </a>
    </div>
  </div>
</template>

<script lang="ts" setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'
import JsBarcode from 'jsbarcode'
import JSZip from 'jszip'
import html2canvas from 'html2canvas'

const file = ref<File | null>(null)
const userNumber = ref<number | null>(null)
const errorMessage = ref<string | null>(null)
const isGenerating = ref<boolean>(false)
const downloadLink = ref<string | null>(null)

const handleFileUpload = (event: Event) => {
  const target = event.target as HTMLInputElement
  if (target.files && target.files[0]) {
    file.value = target.files[0]
  }
}

const generateLabels = async () => {
  errorMessage.value = null
  downloadLink.value = null

  if (!file.value) {
    errorMessage.value = 'Veuillez télécharger un fichier Excel.'
    return
  }

  if (userNumber.value === null || userNumber.value === undefined) {
    errorMessage.value = 'Veuillez entrer un numéro.'
    return
  }

  isGenerating.value = true

  try {
    const data = await file.value.arrayBuffer()
    const workbook = XLSX.read(data)
    const firstSheetName = workbook.SheetNames[0]
    const firstSheet = workbook.Sheets[firstSheetName]
    const jsonData: any[] = XLSX.utils.sheet_to_json(firstSheet, { defval: '' })

    // Validate columns
    const requiredColumns = ['Nom', 'Valeur Option1', 'Prix']
    const header = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })[0]
    const missingColumns = requiredColumns.filter(col => !header.includes(col))
    if (missingColumns.length > 0) {
      errorMessage.value = `Colonnes manquantes : ${missingColumns.join(', ')}`
      isGenerating.value = false
      return
    }

    // Determine if 'Code-barres' exists
    let codeBarresColIndex = header.indexOf('Code-barres')
    if (codeBarresColIndex === -1) {
      // Add 'Code-barres' as the next column
      codeBarresColIndex = header.length
      header.push('Code-barres')
      XLSX.utils.sheet_add_aoa(firstSheet, [header], { origin: 'A1' })
    }

    // Function to format date as DDMMYY
    const formatDate = (date: Date): string => {
      const dd = String(date.getDate()).padStart(2, '0')
      const mm = String(date.getMonth() + 1).padStart(2, '0') // Months are zero-based
      const yy = String(date.getFullYear()).slice(-2)
      return `${dd}${mm}${yy}`
    }

    const today = new Date()
    const dateStr = formatDate(today) // DDMMYY

    const zip = new JSZip()

    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i]
      const identifier = `SELEC${dateStr}${userNumber.value}${i}`
      row['Code-barres'] = identifier

      // Update the 'Code-barres' cell in the worksheet
      const cellAddress = XLSX.utils.encode_cell({
        c: codeBarresColIndex,
        r: i + 1, // +1 because sheet rows are 0-indexed and first row is header
      })
      firstSheet[cellAddress] = { v: identifier }

      // Create label
      const labelDataURL = await createLabel(row['Nom'], row['Valeur Option1'], row['Prix'], identifier)
      const base64Data = labelDataURL.split(',')[1]
      zip.file(`etiquette_${i}.png`, base64Data, { base64: true })
    }

    // Optionally, add the updated Excel file to the ZIP
    const updatedExcel = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    zip.file('updated_file.xlsx', updatedExcel)

    const zipBlob = await zip.generateAsync({ type: 'blob' })
    downloadLink.value = URL.createObjectURL(zipBlob)
  } catch (error) {
    console.error(error)
    errorMessage.value = 'Une erreur est survenue lors du traitement.'
  } finally {
    isGenerating.value = false
  }
}

const createLabel = async (nom: string, valeurOption1: string, prix: string, identifier: string): Promise<string> => {
  return new Promise((resolve, reject) => {
    // Create a hidden div to render the label
    const labelDiv = document.createElement('div')
    labelDiv.style.width = `${mmToPx(60)}px` // 60mm
    labelDiv.style.height = `${mmToPx(30)}px` // 30mm
    labelDiv.style.position = 'absolute'
    labelDiv.style.left = '-9999px'
    labelDiv.style.top = '-9999px'
    labelDiv.style.display = 'flex'
    labelDiv.style.flexDirection = 'column'
    labelDiv.style.justifyContent = 'center'
    labelDiv.style.alignItems = 'center'
    labelDiv.style.border = '1px solid #000'
    labelDiv.style.boxSizing = 'border-box'
    labelDiv.style.backgroundColor = '#ffffff'
    labelDiv.style.padding = '5px' // Add padding to prevent content from touching edges

    // Add Nom
    const nomElement = document.createElement('div')
    nomElement.style.fontSize = '14px'
    nomElement.style.fontWeight = 'bold'
    nomElement.textContent = nom
    labelDiv.appendChild(nomElement)

    // Add Taille (Valeur Option1)
    const tailleElement = document.createElement('div')
    tailleElement.style.fontSize = '12px'
    tailleElement.textContent = `Taille : ${valeurOption1}`
    labelDiv.appendChild(tailleElement)

    // Add Prix with Euro Symbol
    const prixElement = document.createElement('div')
    prixElement.style.fontSize = '12px'
    prixElement.textContent = `${prix}€`
    labelDiv.appendChild(prixElement)

    // Add spacing between Prix and Barcode
    const spacer = document.createElement('div')
    spacer.style.height = '10px' // Increased spacing to prevent touching
    labelDiv.appendChild(spacer)

    // Add Barcode
    const barcodeCanvas = document.createElement('canvas')
    JsBarcode(barcodeCanvas, identifier, {
      format: 'CODE128',
      displayValue: false,
      width: 1,
      height: 30, // Reduced height for smaller barcode
      margin: 2,  // Added margin for spacing
    })
    labelDiv.appendChild(barcodeCanvas)

    document.body.appendChild(labelDiv)

    // Use html2canvas to convert the div to PNG
    html2canvas(labelDiv, { scale: 2 })
      .then(canvas => {
        const dataURL = canvas.toDataURL('image/png')
        document.body.removeChild(labelDiv)
        resolve(dataURL)
      })
      .catch(err => {
        document.body.removeChild(labelDiv)
        reject(err)
      })
  })
}

const mmToPx = (mm: number, dpi: number = 96): number => {
  return Math.round((mm / 25.4) * dpi)
}
</script>

<style scoped>
/* Optional: Add any additional styles if necessary */
</style>
