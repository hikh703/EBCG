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
        :download="`etiquettes_portant_${userNumber}.zip`"
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

const replaceEmptyWithDash = (value: string): string => {
  return value === "" ? "-" : value
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
    const workbook = XLSX.read(data, { cellDates: true, cellNF: false, cellText: false })
    const firstSheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[firstSheetName]

    const jsonData: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' })

    const headerRow = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0]
    const codeBarresIndex = headerRow.indexOf('Code-barres')
    if (codeBarresIndex === -1) {
      errorMessage.value = 'La colonne "Code-barres" est introuvable.'
      isGenerating.value = false
      return
    }

    const requiredColumns = ['Nom', 'Valeur Option1', 'Prix', 'Code-barres']
    const missingColumns = requiredColumns.filter(col => !headerRow.includes(col))
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

    const zip = new JSZip()

    const codeBarresColLetter = XLSX.utils.encode_col(codeBarresIndex)
    const range = XLSX.utils.decode_range(sheet['!ref'] as string)

    // If row 1 is headers, data starts on row 2 in Excel terms:
    const dataStartRow = range.s.r + 2 // range.s.r is typically 0 if the first row is headers

    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i]
      const identifier = `SELEC${dateStr}${userNumber.value}${i}`

      // Replace empty cells with "-"
      row['Nom'] = replaceEmptyWithDash(row['Nom'])
      row['Valeur Option1'] = replaceEmptyWithDash(row['Valeur Option1'])
      row['Prix'] = replaceEmptyWithDash(row['Prix'])
      row['Valeur Option2'] = replaceEmptyWithDash(row['Valeur Option2'])

      // Write the Code-barres value
      const cellAddress = codeBarresColLetter + (dataStartRow + i)
      sheet[cellAddress] = { t: 's', v: identifier }

      // Write back the modified row to the sheet (excluding 'Code-barres')
      headerRow.forEach((header, colIndex) => {
        if (header === 'Code-barres') return // Skip 'Code-barres' as it's already handled
        const colLetter = XLSX.utils.encode_col(colIndex)
        const currentRow = dataStartRow + i
        const cell = sheet[colLetter + currentRow]
        if (cell) {
          cell.v = row[header]
          cell.t = 's' // Assuming all values are strings
        } else {
          sheet[colLetter + currentRow] = { t: 's', v: row[header] }
        }
      })

      // Determine display values for labels
      const displayNom = row['Nom'] !== '-' ? row['Nom'] : ''
      const displayTaille = row['Valeur Option1'] !== '-' ? `Taille : ${row['Valeur Option1']}` : 'Taille :'
      const displayPrix = row['Prix'] !== '-' ? `${row['Prix']}€` : ''

      // Generate the label image
      const labelDataURL = await createLabel(displayNom, displayTaille, displayPrix, identifier)
      const base64Data = labelDataURL.split(',')[1]

      // Use the product name in the filename, avoid empty filenames
      const productName = row['Nom'] !== '-' ? row['Nom'] : `Produit_${i}`
      zip.file(`etiquette_${productName}.png`, base64Data, { base64: true })
    }

    // Write the updated workbook to Excel
    const updatedExcel = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    // Use userNumber in the Excel file name
    zip.file(`portant_${userNumber.value}.xlsx`, updatedExcel)

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
    const dpi = 300;
    const widthPx = Math.round((60 / 25.4) * dpi);
    const heightPx = Math.round((30 / 25.4) * dpi);

    const topOffset = 0;    
    const bottomOffset = 20;

    const labelDiv = document.createElement('div');
    labelDiv.style.width = `${widthPx}px`;
    labelDiv.style.height = `${heightPx}px`;
    labelDiv.style.position = 'relative';
    labelDiv.style.backgroundColor = '#fff';
    labelDiv.style.overflow = 'hidden';

    let nameFontSize = Math.round((3 / 25.4) * dpi);
    let otherFontSize = Math.round((3 / 25.4) * dpi);

    const topDiv = document.createElement('div');
    topDiv.style.position = 'absolute';
    topDiv.style.top = `${topOffset}px`;
    topDiv.style.left = '0';
    topDiv.style.width = '100%';
    topDiv.style.display = 'flex';
    topDiv.style.flexDirection = 'column';
    topDiv.style.justifyContent = 'flex-start';
    topDiv.style.alignItems = 'center'; 
    topDiv.style.textAlign = 'center';
    topDiv.style.paddingLeft = '10px'; 
    topDiv.style.paddingRight = '10px'; 
    topDiv.style.boxSizing = 'border-box';
    topDiv.style.whiteSpace = 'normal';
    topDiv.style.wordBreak = 'break-word';

    const nomElement = document.createElement('div');
    nomElement.style.fontSize = `${nameFontSize}px`;
    nomElement.style.lineHeight = '1.2em';
    nomElement.style.textAlign = 'center';
    nomElement.textContent = nom;

    const tailleElement = document.createElement('div');
    tailleElement.style.fontSize = `${otherFontSize}px`;
    tailleElement.style.lineHeight = '1.2em';
    tailleElement.style.textAlign = 'center';
    tailleElement.textContent = taille;
    tailleElement.style.paddingTop = '15px'

    const prixElement = document.createElement('div');
    prixElement.style.fontSize = `${otherFontSize}px`;
    prixElement.style.lineHeight = '1.2em';
    prixElement.style.textAlign = 'center';
    prixElement.style.paddingTop = '15px'
    prixElement.textContent = prix;

    topDiv.appendChild(nomElement);
    topDiv.appendChild(tailleElement);
    topDiv.appendChild(prixElement);
    labelDiv.appendChild(topDiv);

    const bottomDiv = document.createElement('div');
    bottomDiv.style.position = 'absolute';
    bottomDiv.style.bottom = `${bottomOffset}px`;
    bottomDiv.style.left = '0';
    bottomDiv.style.width = '100%';
    bottomDiv.style.display = 'flex';
    bottomDiv.style.justifyContent = 'center'; 
    bottomDiv.style.alignItems = 'flex-end';
    bottomDiv.style.boxSizing = 'border-box';

    const barcodeCanvas = document.createElement('canvas');
    JsBarcode(barcodeCanvas, identifier, {
      format: 'CODE128',
      displayValue: false,
      width: Math.round((0.25 / 25.4) * dpi),
      height: Math.round((10 / 25.4) * dpi),
      margin: 0,
    });
    bottomDiv.appendChild(barcodeCanvas);
    labelDiv.appendChild(bottomDiv);

    document.body.appendChild(labelDiv);

    const fitText = () => {
      const lines = Math.ceil(nomElement.scrollHeight / parseFloat(window.getComputedStyle(nomElement).lineHeight));
      const maxLines = 2;
      if (lines > maxLines && nameFontSize > 5) {
        nameFontSize -= 1;
        otherFontSize -= 1;
        nomElement.style.fontSize = `${nameFontSize}px`;
        tailleElement.style.fontSize = `${otherFontSize}px`;
        prixElement.style.fontSize = `${otherFontSize}px`;
        fitText();
      }
    };

    setTimeout(() => {
      fitText();
      html2canvas(labelDiv, { scale: 1 }).then(canvas => {
        const dataURL = canvas.toDataURL('image/png');
        document.body.removeChild(labelDiv);
        resolve(dataURL);
      }).catch(err => {
        document.body.removeChild(labelDiv);
        reject(err);
      });
    }, 0);
  });
}
</script>
