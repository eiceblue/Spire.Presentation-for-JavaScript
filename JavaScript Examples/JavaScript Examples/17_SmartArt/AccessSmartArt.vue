<template>
  <span>Click the following button to access SmartArt and get SmartArt node parameters. </span>
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
        strB.push('Access SmartArt nodes.');
        strB.push('Here is the SmartArt node parameters details:');
        let outString = '';
        let node;
        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
          let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
          if (shape instanceof wasmModule.ISmartArt) {
            //Get the SmartArt and collect nodes
            let sa = shape;
            let nodes = sa.Nodes;

            //Traverse through all nodes inside SmartArt
            for (let j = 0; j < nodes.Count; j++) {
              //Access SmartArt node at index i
              node = nodes.get_Item(j);
              //Print the SmartArt node parameters
              outString = `Node text = ${node.TextFrame.Text}, Node level = ${node.Level}, Node Position = ${node.Position}`;
              strB.push(outString);
            }
          }
        }

        // Define the output file name
        const outputFileName = 'AccessSmartArt.txt';

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
