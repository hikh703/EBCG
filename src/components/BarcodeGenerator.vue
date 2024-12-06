<template>
    <div class="container mx-auto p-4">
      <h1 class="text-2xl font-bold mb-4">Excel to Barcode Generator</h1>
      <input type="file" @change="handleFile" accept=".xlsx, .xls" class="mb-4"/>
      <div v-if="barcodes.length">
        <div v-for="(code, index) in barcodes" :key="index" class="mb-4">
          <svg :id="'barcode' + index"></svg>
          <p class="text-center mt-2">{{ code }}</p>
        </div>
      </div>
    </div>
  </template>
  
  <script>
  import * as XLSX from 'xlsx';
  import JsBarcode from 'jsbarcode';
  
  export default {
    name: 'BarcodeGenerator',
    data() {
      return {
        barcodes: [],
      };
    },
    methods: {
      handleFile(event) {
        const file = event.target.files[0];
        const reader = new FileReader();
        reader.onload = (e) => {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheet];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          // Assuming barcodes are in the first column
          this.barcodes = json.slice(1).map(row => row[0]); // Skip header
          this.$nextTick(() => {
            this.generateBarcodes();
          });
        };
        reader.readAsArrayBuffer(file);
      },
      generateBarcodes() {
        this.barcodes.forEach((code, index) => {
          JsBarcode(`#barcode${index}`, code, {
            format: "CODE128",
            displayValue: true,
            width: 2,
            height: 40,
            margin: 10,
          });
        });
      },
    },
  };
  </script>
  
  <style scoped>
  /* Add any component-specific styles here */
  </style>
  