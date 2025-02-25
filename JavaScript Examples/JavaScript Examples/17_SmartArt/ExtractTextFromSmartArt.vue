<template>
  <span>Click the following button to extract text from SmartArt in PowerPoint file. </span>
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
        let inputFileName = 'ExtractTextFromSmartArt.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        //Traverse through all the slides of the PPT file and find the SmartArt shapes.
        let st = [];
        st.push('Below is extracted text from SmartArt:');
        for (let i = 0; i < ppt.Slides.Count; i++) {
          for (let j = 0; j < ppt.Slides.get_Item(i).Shapes.Count; j++) {
            if (ppt.Slides.get_Item(i).Shapes.get_Item(j) instanceof wasmModule.ISmartArt) {
              let smartArt = ppt.Slides.get_Item(i).Shapes.get_Item(j);

              //Extract text from SmartArt and append to the StringBuilder object.
              for (let k = 0; k < smartArt.Nodes.Count; k++) {
                st.push(smartArt.Nodes.get_Item(k).TextFrame.Text);
              }
            }
          }
        }

        // Define the output file name
        const outputFileName = 'ExtractTextFromSmartArt.txt';

        // Save result file
        wasmModule.FS.writeFile(outputFileName, st.join('\r\n'));

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
