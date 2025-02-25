<template>
  <span>Click the following button to access SmartArt layout. </span>
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
        let inputFileName = 'SmartArt.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        let strB = [];

        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
          let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
          if (shape instanceof wasmModule.ISmartArt) {
            //Get the SmartArt and collect nodes
            let sa = shape;
            //Check SmartArt Layout
            let layout = sa.LayoutType.toString();
            strB.push('SmartArt layout type is ' + layout);
          }
        }

        // Define the output file name
        const outputFileName = 'AccessSmartArtLayout.txt';

        // Save result file
        wasmModule.FS.writeFile(outputFileName, strB.join('\r\n'));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'text/plain',
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
