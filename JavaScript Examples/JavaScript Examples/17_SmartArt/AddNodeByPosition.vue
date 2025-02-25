<template>
  <span>Click the following button to add SmartArt node by position. </span>
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
        let inputFileName = 'AddSmartArtNode2.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
          let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
          if (shape instanceof wasmModule.ISmartArt) {
            //Get the SmartArt and collect nodes
            let smartArt = shape;
            let position = 0;
            //Add a new node at specific position
            let node = smartArt.Nodes.AddNodeByPosition(position);
            //Add text and set the text style
            node.TextFrame.Text = 'New Node';
            node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Red;

            //Get a node
            node = smartArt.Nodes.get_Item(1);
            position = 1;
            //Add a new child node at specific position
            let childNode = node.ChildNodes.AddNodeByPosition(position);
            //Add text and set the text style
            node.TextFrame.Text = 'New child node';
            node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Blue;
          }
        }

        // Define the output file name
        const outputFileName = 'AddNodeByPosition.pptx';

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
