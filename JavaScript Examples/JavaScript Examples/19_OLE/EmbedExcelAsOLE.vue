<template>
  <span>Click the following button to embed an Excel as an OLE object. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'EmbedExcelAsOLE.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        // Load the sample png into the virtual file system (VFS)
        let imageFileName = 'EmbedExcelAsOLE.png';
        await wasmModule.FetchFileToVFS(imageFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        let stream = wasmModule.Stream.CreateByFile(imageFileName);
        let oleImage = ppt.Images.Append({ stream });
        stream.Close();

        let rec = wasmModule.RectangleF.FromLTRB(80, 60, 550, 450);
        //Insert an OLE object to presentation based on the Excel data
        let objectData = wasmModule.Stream.CreateByFile(inputFileName);
        let oleObject = ppt.Slides.get_Item(0).Shapes._AppendOleObject('excel', objectData, rec);
        oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;

        oleObject.ProgId = 'Excel.Sheet.12';
        // Define the output file name
        const outputFileName = 'EmbedExcelAsOLE.pptx';

        // Save result file
        ppt.SaveToFile({
          file: outputFileName,
          fileFormat: wasmModule.FileFormat.Pptx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
