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
    const jsonData: any[] = XLSX.utils.sheet_to_json(firstSheet, { defval: '', blankrows: true })

    // Validate columns
    const requiredColumns = ['Nom', 'Valeur Option1', 'Prix']
    const header = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })[0]
    const missingColumns = requiredColumns.filter(col => !header.includes(col))
    if (missingColumns.length > 0) {
      errorMessage.value = `Colonnes manquantes : ${missingColumns.join(', ')}`
      isGenerating.value = false
      return
    }

    // Ensure all rows are processed
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

    const updatedData = jsonData.map((row, index) => {
      const identifier = `SELEC${dateStr}${userNumber.value}${index}`
      row['Code-barres'] = identifier
      return row
    })

    const updatedSheet = XLSX.utils.json_to_sheet(updatedData, { skipHeader: false })
    Object.keys(updatedSheet).forEach(cell => {
      if (!cell.startsWith('!')) {
        updatedSheet[cell].z = '@'
        updatedSheet[cell].t = 's' // Explicitly set as text
        if (typeof updatedSheet[cell].v === 'number') {
          updatedSheet[cell].v = String(updatedSheet[cell].v)
        }
      }
    })

    workbook.Sheets[firstSheetName] = updatedSheet

    const zip = new JSZip()
    for (let i = 0; i < updatedData.length; i++) {
      const row = updatedData[i]
      const labelDataURL = await createLabel(row['Nom'], row['Valeur Option1'], row['Prix'], row['Code-barres'])
      const base64Data = labelDataURL.split(',')[1]
      zip.file(`etiquette_${i}.png`, base64Data, { base64: true })
    }

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
    const dpi = 300
    const widthPx = Math.round((60 / 25.4) * dpi) // Convert 60 mm to pixels
    const heightPx = Math.round((30 / 25.4) * dpi) // Convert 30 mm to pixels

    const labelDiv = document.createElement('div')
    labelDiv.style.width = `${widthPx}px`
    labelDiv.style.height = `${heightPx}px`
    labelDiv.style.position = 'absolute'
    labelDiv.style.left = '-9999px'
    labelDiv.style.top = '-9999px'
    labelDiv.style.display = 'flex'
    labelDiv.style.flexDirection = 'column'
    labelDiv.style.justifyContent = 'space-between' // Ensures spacing between items
    labelDiv.style.alignItems = 'center'
    labelDiv.style.border = '1px solid #000'
    labelDiv.style.backgroundColor = '#fff'
    labelDiv.style.padding = `${Math.round((2 / 25.4) * dpi)}px`

    // Add Nom
    const nomElement = document.createElement('div')
    nomElement.style.fontSize = `${Math.round((3 / 25.4) * dpi)}px`
    nomElement.style.fontWeight = 'bold'
    nomElement.textContent = nom
    labelDiv.appendChild(nomElement)

    // Add Taille (Valeur Option1)
    const tailleElement = document.createElement('div')
    tailleElement.style.fontSize = `${Math.round((2.5 / 25.4) * dpi)}px`
    tailleElement.textContent = `Taille : ${valeurOption1}`
    labelDiv.appendChild(tailleElement)

    // Add Prix with Euro Symbol
    const prixElement = document.createElement('div')
    prixElement.style.fontSize = `${Math.round((2.5 / 25.4) * dpi)}px`
    prixElement.textContent = `${prix}€`
    labelDiv.appendChild(prixElement)

    // Spacer to push the barcode to the bottom
    const spacer = document.createElement('div')
    spacer.style.flexGrow = '1' // Pushes the barcode to the bottom
    spacer.style.padding = '5px'
    labelDiv.appendChild(spacer)

    // Add Barcode
    const barcodeCanvas = document.createElement('canvas')
    JsBarcode(barcodeCanvas, identifier, {
      format: 'CODE128',
      displayValue: false,
      width: Math.round((0.25 / 25.4) * dpi),
      height: Math.round((8 / 25.4) * dpi),
      margin: 4,
    })
    labelDiv.appendChild(barcodeCanvas)

    document.body.appendChild(labelDiv)

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
