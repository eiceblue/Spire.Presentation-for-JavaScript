<template>
  <span>Click the following button to find the first matched text and modify its style in a PPT document.</span>
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
        let inputFileName = "TextTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Find first "Spire"
        let text = "Spire";
        let textRange = ppt.Slides.get_Item(0).FindFirstTextAsRange(text);

        // Modify the style
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = wasmModule.Color.get_Red();
        textRange.FontHeight = 28;
        textRange.LatinFont = wasmModule.TextFont.Create("Calibri");
        textRange.IsBold = wasmModule.TriState.True;
        textRange.IsItalic = wasmModule.TriState.True;
        textRange.TextUnderlineType = wasmModule.TextUnderlineType.Double;
        textRange.TextStrikethroughType = wasmModule.TextStrikethroughType.Single;

        const outputFileName = "ModifyStyleOfFirstFoundText.pptx";

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
