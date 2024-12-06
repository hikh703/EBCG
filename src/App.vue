<template>
  <div id="app">
    <h1>Excel to Etiquette Generator</h1>
    <input type="file" @change="handleFileUpload" accept=".xlsx, .xls" />
    <button @click="generateEtiquettes" :disabled="!data.length">Generate Etiquettes</button>
    <div v-if="isGenerating">
      <p>Generating images and preparing ZIP...</p>
    </div>
    <!-- Hidden container to render etikettes for image capture -->
    <div ref="etiquetteContainer" style="display: none;">
      <div
        v-for="(item, index) in data"
        :key="index"
        class="etiquette"
        :id="'etiquette-' + index"
        style="width: 300px; padding: 20px; border: 1px solid #000; margin-bottom: 10px; text-align: center;"
      >
        <h2>{{ item.ProductName }}</h2>
        <svg :id="'barcode-' + index"></svg>
      </div>
    </div>
  </div>
</template>

<script>
import * as XLSX from 'xlsx';
import JsBarcode from 'jsbarcode';
import html2canvas from 'html2canvas';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

export default {
  name: 'App',
  data() {
    return {
      data: [],
      isGenerating: false,
    };
  },
  methods: {
    handleFileUpload(event) {
      const file = event.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const workbook = XLSX.read(binaryStr, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Assuming the first row is the header
        const headers = jsonData[0];
        const productNameIndex = headers.indexOf('ProductName'); // Adjust based on your Excel header

        if (productNameIndex === -1) {
          alert('Please ensure there is a "ProductName" column in the Excel file.');
          return;
        }

        // Extract data starting from the second row
        this.data = jsonData.slice(1).map((row, index) => ({
          id: index + 1, // Unique identifier
          ProductName: row[productNameIndex],
        }));
      };
      reader.readAsBinaryString(file);
    },
    async generateEtiquettes() {
      if (!this.data.length) {
        alert('No data to generate etikettes.');
        return;
      }

      this.isGenerating = true;

      const zip = new JSZip();
      const promises = this.data.map(async (item, index) => {
        // Generate barcode
        const barcodeSvg = document.getElementById(`barcode-${index}`);
        if (!barcodeSvg) return;

        JsBarcode(barcodeSvg, item.id.toString(), {
          format: 'CODE128',
          width: 2,
          height: 50,
          displayValue: false,
        });

        // Wait for the barcode to render
        await this.$nextTick();

        // Capture the etikette as PNG
        const etiquetteElement = document.getElementById(`etiquette-${index}`);
        if (!etiquetteElement) return;

        const canvas = await html2canvas(etiquetteElement, { backgroundColor: '#fff' });
        const imgData = canvas.toDataURL('image/png');

        // Add to ZIP
        const base64Data = imgData.split(',')[1];
        zip.file(`etiquette_${item.id}.png`, base64Data, { base64: true });
      });

      // Wait for all etikettes to be processed
      await Promise.all(promises);

      // Generate ZIP
      const zipContent = await zip.generateAsync({ type: 'blob' });

      // Trigger download
      saveAs(zipContent, 'etiquettes.zip');

      this.isGenerating = false;
    },
  },
};
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  padding: 20px;
}

button {
  margin-top: 10px;
  padding: 10px 20px;
}

.etiquette {
  /* Styles for the etiquette, adjust as needed */
}
</style>
