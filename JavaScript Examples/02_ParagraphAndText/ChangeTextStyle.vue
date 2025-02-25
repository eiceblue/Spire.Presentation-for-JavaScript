<template>
  <span>Click the following button to change the font and color of text in a PPT document.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "ChangeTextStyle.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the first shape from the first slide
        let shape = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        let paras = shape.TextFrame.Paragraphs;

        // Set the style for the text content in the first paragraph
        for(let i = 0;i < paras.get_Item(0).TextRanges.Count;i++){
            paras.get_Item(0).TextRanges.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
            paras.get_Item(0).TextRanges.get_Item(i).Fill.SolidColor.Color = wasmModule.Color.get_ForestGreen();
            paras.get_Item(0).TextRanges.get_Item(i).LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
            paras.get_Item(0).TextRanges.get_Item(i).FontHeight = 14;
        }

        // Set the style for the text content in the third paragraph
        for(let i = 0;i < paras.get_Item(2).TextRanges.Count;i++){
            paras.get_Item(2).TextRanges.get_Item(i).Fill.FillType = wasmModule.FillFormatType.Solid;
            paras.get_Item(2).TextRanges.get_Item(i).Fill.SolidColor.Color = wasmModule.Color.get_CornflowerBlue();
            paras.get_Item(2).TextRanges.get_Item(i).LatinFont = wasmModule.TextFont.Create("Calibri");
            paras.get_Item(2).TextRanges.get_Item(i).FontHeight = 16;
            paras.get_Item(2).TextRanges.get_Item(i).TextUnderlineType = wasmModule.TextUnderlineType.Dashed;
        }

        const outputFileName = "ChangeTextStyle.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

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
