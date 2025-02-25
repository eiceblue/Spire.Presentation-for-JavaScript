<template>
  <span>Click the following button to create a SmartArt shape. </span>
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

        let sa = ppt.Slides.get_Item(0).Shapes.AppendSmartArt(200, 60, 300, 300, wasmModule.SmartArtLayoutType.Gear);

        //Set type and color of smartart
        sa.Style = wasmModule.SmartArtStyleType.SubtleEffect;
        sa.ColorStyle = wasmModule.SmartArtColorType.GradientLoopAccent3;

        //Remove all shapes
        for (let i = sa.Nodes.Count - 1; i >= 0; i--) {
          sa.Nodes.RemoveNode({ index: i });
        }

        //Add two custom shapes with text
        let node = sa.Nodes.AddNode();
        sa.Nodes.get_Item(0).TextFrame.Text = 'aa';
        node = sa.Nodes.AddNode();
        node.TextFrame.Text = 'bb';
        node.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        node.TextFrame.TextRange.Fill.SolidColor.KnownColor = wasmModule.KnownColors.Black;

        // Define the output file name
        const outputFileName = 'CreateSmartArtShape.pptx';

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
