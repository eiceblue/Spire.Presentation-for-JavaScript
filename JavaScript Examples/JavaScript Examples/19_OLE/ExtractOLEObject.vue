<template>
  <span>Click the following button to extract OLE object from PPT document. </span>
  <el-button @click="startProcessing">Start</el-button>

  <a v-if="downloadUrlDoc" :href="downloadUrlDoc" :download="downloadNameDoc"> Click here to download the generated Doc file </a>
  <a v-if="downloadUrlDocx" :href="downloadUrlDocx" :download="downloadNameDocx"> Click here to download the generated Docx file </a>
  <a v-if="downloadUrlXls" :href="downloadUrlXls" :download="downloadNameXls"> Click here to download the generated Xls file </a>
  <a v-if="downloadUrlXlsx" :href="downloadUrlXlsx" :download="downloadNameXlsx"> Click here to download the generated Xlsx file </a>
  <a v-if="downloadUrlPpt" :href="downloadUrlPpt" :download="downloadNamePpt"> Click here to download the generated Ppt file </a>
  <a v-if="downloadUrlPptx" :href="downloadUrlPptx" :download="downloadNamePptx"> Click here to download the generated Pptx file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrlDoc = ref(null);
    const downloadNameDoc = ref('');

    const downloadUrlDocx = ref(null);
    const downloadNameDocx = ref('');

    const downloadUrlXls = ref(null);
    const downloadNameXls = ref('');

    const downloadUrlXlsx = ref(null);
    const downloadNameXlsx = ref('');

    const downloadUrlPpt = ref(null);
    const downloadNamePpt = ref('');

    const downloadUrlPptx = ref(null);
    const downloadNamePptx = ref('');

    const downloadOle = (outputFileName, bytes) => {
      wasmModule.FS.writeFile(outputFileName, bytes);
      const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
      const modifiedFile = new Blob([modifiedFileArray], {});
      switch (outputFileName) {
        // Download the file
        case 'ExtractOLEObject.pptx':
          downloadNamePptx.value = outputFileName;
          downloadUrlPptx.value = URL.createObjectURL(modifiedFile);
          break;
        case 'ExtractOLEObject.ppt':
          downloadNamePpt.value = outputFileName;
          downloadUrlPpt.value = URL.createObjectURL(modifiedFile);
          break;
        case 'ExtractOLEObject.xls':
          downloadNameXls.value = outputFileName;
          downloadUrlXls.value = URL.createObjectURL(modifiedFile);
          break;
        case 'ExtractOLEObject.xlsx':
          downloadNameXlsx.value = outputFileName;
          downloadUrlXlsx.value = URL.createObjectURL(modifiedFile);
          break;
        case 'ExtractOLEObject.doc':
          downloadNameDoc.value = outputFileName;
          downloadUrlDoc.value = URL.createObjectURL(modifiedFile);
          break;
        case 'ExtractOLEObject.docx':
          downloadNameDocx.value = outputFileName;
          downloadUrlDocx.value = URL.createObjectURL(modifiedFile);
          break;
      }
    };

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'ExtractOLEObject.pptx';
        // Define the name of the output file
        const outFileName_pptx = 'ExtractOLEObject.pptx';
        const outFileName_ppt = 'ExtractOLEObject.ppt';
        const outFileName_xls = 'ExtractOLEObject.xls';
        const outFileName_xlsx = 'ExtractOLEObject.xlsx';
        const outFileName_doc = 'ExtractOLEObject.doc';
        const outFileName_docx = 'ExtractOLEObject.docx';

        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        //Loop through the slides and shapes
        for (let i = 0; i < ppt.Slides.Count; i++) {
          let slide = ppt.Slides.get_Item(i);
          for (let j = 0; j < slide.Shapes.Count; j++) {
            let shape = slide.Shapes.get_Item(j);
            if (shape instanceof wasmModule.IOleObject) {
              //Find OLE object
              let oleObject = shape;

              //Get its data and write to file
              let bytes = oleObject.Data;
              switch (oleObject.ProgId) {
                case 'Excel.Sheet.8':
                  downloadOle(outFileName_xls, bytes);
                  break;
                case 'Excel.Sheet.12':
                  downloadOle(outFileName_xlsx, bytes);
                  break;
                case 'Word.Document.8':
                  downloadOle(outFileName_doc, bytes);
                  break;
                case 'Word.Document.12':
                  downloadOle(outFileName_docx, bytes);
                  break;
                case 'PowerPoint.Show.8':
                  downloadOle(outFileName_ppt, bytes);
                  break;
                case 'PowerPoint.Show.12':
                  downloadOle(outFileName_pptx, bytes);
                  break;
              }
            }
          }
        }

        // Clean up resources
        ppt.Dispose();
      }
    };

    return {
      startProcessing,
      downloadNamePptx,
      downloadUrlPptx,
      downloadNamePpt,
      downloadUrlPpt,
      downloadNameXls,
      downloadUrlXls,
      downloadNameXlsx,
      downloadUrlXlsx,
      downloadNameDoc,
      downloadUrlDoc,
      downloadNameDocx,
      downloadUrlDocx,
    };
  },
};
</script>
