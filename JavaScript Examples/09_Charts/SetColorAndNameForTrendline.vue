<template>
  <span>Click the following button to set color and name for trendline in chart.</span>
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
        let inputFileName = "SetColorAndNameForTrendline.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);


        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);
        //Find the first chart in the first Slide
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Find the first trendline in the chart
        let trendline = chart.Series.get_Item(0).TrendLines[0];

        //Set name for trendline
        trendline.Name = "trendlineName";

        //Set color for trendline
        trendline.Line.FillType = wasmModule.FillFormatType.Solid;
        trendline.Line.SolidFillColor.Color = wasmModule.Color.get_Red();

        // Define the output file name
        const outputFileName = "SetColorAndNameForTrendline_out.pptx";

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
