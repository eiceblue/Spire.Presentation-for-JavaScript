<template>
  <span>Click the following button to access specific SmartArt child node. </span>
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
        strB.push('Access SmartArt child node at specific position.');
        strB.push('Here is the SmartArt child node parameters details:');
        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
          let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
          if (shape instanceof wasmModule.ISmartArt) {
            //Get the SmartArt and collect nodes
            let sa = shape;
            //Get SmartArt node collection
            let nodes = sa.Nodes;

            //Access SmartArt node at index 0
            let node = nodes.get_Item(0);

            //Access SmartArt child node at index 1
            let childNode = node.ChildNodes.get_Item(1);

            //Print the SmartArt child node parameters
            let outString = `Node text = ${childNode.TextFrame.Text}, Node level = ${childNode.Level}, Node Position = ${childNode.Position}`;

            strB.push(outString);
          }
        }

        // Define the output file name
        const outputFileName = 'AccessSpecificChildNode.txt';

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
