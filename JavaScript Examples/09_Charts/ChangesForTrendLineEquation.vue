<template>
  <span>Click the following button to change font size and position for TrendLine equation. </span>
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
        let inputFileName = 'TrendlineEquation.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create Presentation
        let ppt = wasmModule.Presentation.Create();

        //Load ppt file
        ppt.LoadFromFile(inputFileName);

        //Get chart on the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Get the first trendline
        let trendline = chart.Series.get_Item(0).TrendLines[0];

        //Change font size for trendline Equation text
        for (let i = 0; i < trendline.TrendLineLabel.TextFrameProperties.Paragraphs.Count; i++) {
          let para = trendline.TrendLineLabel.TextFrameProperties.Paragraphs.get_Item(i);
          para.DefaultCharacterProperties.FontHeight = 20;
          for (let j = 0; j < para.TextRanges.Count; j++) {
            let range = para.TextRanges.get_Item(j);
            range.FontHeight = 20;
          }
        }

        //Change position for trendline Equation
        trendline.TrendLineLabel.OffsetX = -0.1;
        trendline.TrendLineLabel.OffsetY = -0.05;

        // Define the output file name
        const outputFileName = "ChangesForTrendLineEquation_out.pptx";

        // Save the document to the specified path
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

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
