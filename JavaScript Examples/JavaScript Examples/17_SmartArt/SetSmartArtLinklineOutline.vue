<template>
  <span>Click the following button to set SmartArt linkline outline. </span>
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
        let inputFileName = 'CreateSmartArtShape.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT
        ppt.LoadFromFile(inputFileName);

        let smartArt = ppt.Slides.get_Item(0).Shapes.get_Item(0);
        let count = smartArt.Nodes.Count;
        let node;
        //Loop through all smartArts
        for (let i = 0; i < count; i++) {
          node = smartArt.Nodes.get_Item(i);
          //Set the line type
          node.LinkLine.FillType = wasmModule.FillFormatType.Solid;
          //Set the line color
          node.LinkLine.SolidFillColor.Color = wasmModule.Color.get_Red();
          //Set the line width
          node.LinkLine.Width = 2;
          //Set the line DashStyle
          node.LinkLine.DashStyle = wasmModule.LineDashStyleType.SystemDash;
        }

        // Define the output file name
        const outputFileName = 'SetSmartArtLinklineOutline.pptx';

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
