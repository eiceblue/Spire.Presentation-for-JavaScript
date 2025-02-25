<template>
  <span>Click the following button to embed Zip in PowerPoint. </span>
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
        let inputFileName = 'EmbedZipIntoPPT.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        // Load the zip file into the virtual file system (VFS)
        let inputFile_zName = 'test.zip';
        await wasmModule.FetchFileToVFS(inputFile_zName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Load the sample png into the virtual file system (VFS)
        let inputFile_IName = 'icon.png';
        await wasmModule.FetchFileToVFS(inputFile_IName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();
        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        //Load a zip object
        let data = wasmModule.Stream.CreateByFile(inputFile_zName);

        let rec = wasmModule.RectangleF.FromLTRB(80, 60, 180, 160);

        //Insert the zip object to presentation
        let ole = ppt.Slides.get_Item(0).Shapes._AppendOleObjectOOR(inputFile_zName, data, rec);
        ole.ProgId = 'Package';
        let stream = wasmModule.Stream.CreateByFile(inputFile_IName);
        let oleImage = ppt.Images.Append({ stream: stream });

        ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
        // Define the output file name
        const outputFileName = 'EmbedZipIntoPPT.pptx';

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
