<template>
  <div class="p-6">
    <h1 class="text-2xl font-bold mb-4">Générateur d'Étiquettes</h1>
    <div class="mb-4">
      <label class="block text-gray-700">Télécharger le fichier Excel :</label>
      <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" class="mt-2" />
    </div>
    <div class="mb-4">
      <label class="block text-gray-700">Entrez un numéro de portant :</label>
      <input
        type="number"
        v-model.number="userNumber"
        class="mt-2 p-2 border rounded w-full"
        :class="{ 'border-red-500': errorMessage }"
      />
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
    errorMessage.value = null
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
    errorMessage.value = 'Entrez un numéro de portant.'
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

    if (jsonData.length === 0) {
      errorMessage.value = 'Le fichier Excel ne contient aucune ligne.'
      isGenerating.value = false
      return
    }

    const today = new Date()
    const formatDate = (date: Date): string => {
      const dd = String(date.getDate()).padStart(2, '0')
      const mm = String(date.getMonth() + 1).padStart(2, '0')
      const yy = String(date.getFullYear()).slice(-2)
      return `${dd}${mm}${yy}`
    }
    const dateStr = formatDate(today)

    const zettleData = jsonData.map((row, index) => {
      const identifier = `SELEC${dateStr}${userNumber.value}${index}`
      return {
        'Product Name': row['Nom'], // Zettle-Compatible Column
        'Price (incl. VAT)': row['Prix'], // Zettle-Compatible Column
        'Barcode': identifier, // Generated Barcode
      }
    })

    // Create Zettle-Compatible Sheet
    const zettleSheet = XLSX.utils.json_to_sheet(zettleData)

    const newWorkbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(newWorkbook, zettleSheet, 'Products')

    const zip = new JSZip()

    // Save Zettle-Compatible Excel File
    const updatedExcel = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' })
    zip.file('zettle_file.xlsx', updatedExcel)

    // Generate Labels
    for (let i = 0; i < zettleData.length; i++) {
      const row = zettleData[i]
      const labelDataURL = await createLabel(row['Product Name'], 'Taille : ' + jsonData[i]['Valeur Option1'], row['Price (incl. VAT)'], row['Barcode'])
      const base64Data = labelDataURL.split(',')[1]
      zip.file(`etiquette_${i}.png`, base64Data, { base64: true })
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' })
    downloadLink.value = URL.createObjectURL(zipBlob)
  } catch (error) {
    console.error(error)
    errorMessage.value = 'Une erreur est survenue lors du traitement.'
  } finally {
    isGenerating.value = false
  }
}

const createLabel = async (nom: string, taille: string, prix: string, identifier: string): Promise<string> => {
  return new Promise((resolve, reject) => {
    const dpi = 300 // High-resolution for printing
    const widthPx = Math.round((60 / 25.4) * dpi) // Convert 60 mm to pixels
    const heightPx = Math.round((30 / 25.4) * dpi) // Convert 30 mm to pixels

    const labelDiv = document.createElement('div')
    labelDiv.style.width = `${widthPx}px`
    labelDiv.style.height = `${heightPx}px`
    labelDiv.style.display = 'flex'
    labelDiv.style.flexDirection = 'column'
    labelDiv.style.justifyContent = 'space-between'
    labelDiv.style.alignItems = 'center'
    labelDiv.style.backgroundColor = '#fff' // Ensure white background
    labelDiv.style.padding = '0'

    // Top Section for Text
    const topDiv = document.createElement('div')
    topDiv.style.display = 'flex'
    topDiv.style.flexDirection = 'column'
    topDiv.style.alignItems = 'center'
    topDiv.style.justifyContent = 'flex-start' // Align closer to the top
    topDiv.style.height = '60%' // Decrease height of top section
    topDiv.style.marginTop = '0' // Ensure no unnecessary margin

    // Add Nom (Name)
    const nomElement = document.createElement('div')
    nomElement.style.fontSize = `${Math.round((4 / 25.4) * dpi)}px`
    nomElement.style.fontWeight = 'bold'
    nomElement.textContent = nom
    topDiv.appendChild(nomElement)

    // Add Taille (Size)
    const tailleElement = document.createElement('div')
    tailleElement.style.fontSize = `${Math.round((3 / 25.4) * dpi)}px`
    tailleElement.textContent = taille
    topDiv.appendChild(tailleElement)

    // Add Prix (Price)
    const prixElement = document.createElement('div')
    prixElement.style.fontSize = `${Math.round((3 / 25.4) * dpi)}px`
    prixElement.textContent = `${prix}€`
    topDiv.appendChild(prixElement)

    labelDiv.appendChild(topDiv)

    // Bottom Section for Barcode
    const bottomDiv = document.createElement('div')
    bottomDiv.style.display = 'flex'
    bottomDiv.style.justifyContent = 'center'
    bottomDiv.style.alignItems = 'flex-end' // Align closer to the bottom
    bottomDiv.style.height = '40%' // Increase height for barcode section
    bottomDiv.style.marginBottom = `${Math.round((2 / 25.4) * dpi)}px` // Space below the barcode

    const barcodeCanvas = document.createElement('canvas')
    JsBarcode(barcodeCanvas, identifier, {
      format: 'CODE128',
      displayValue: false,
      width: Math.round((0.25 / 25.4) * dpi), // Proper scaling for width
      height: Math.round((10 / 25.4) * dpi), // Proper scaling for height
      margin: 0,
    })
    bottomDiv.appendChild(barcodeCanvas)

    labelDiv.appendChild(bottomDiv)

    document.body.appendChild(labelDiv)

    // Render the label into a PNG image
    html2canvas(labelDiv, { scale: 1 }).then(canvas => {
      const dataURL = canvas.toDataURL('image/png')
      document.body.removeChild(labelDiv)
      resolve(dataURL)
    }).catch(err => {
      document.body.removeChild(labelDiv)
      reject(err)
    })
  })
}





</script>
