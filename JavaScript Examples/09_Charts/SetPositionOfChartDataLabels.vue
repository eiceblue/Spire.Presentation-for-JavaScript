<template>
  <span>The following example demonstrates how to set position of chart data labels</span>
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
        
        // Load the input file into the virtual file system (VFS)
        const inputFileName = "Template_Ppt_2.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PowerPoint document
        let ppt = wasmModule.Presentation.Create();

        // Load file from VFS
        ppt.LoadFromFile(inputFileName);

        //Get the chart.
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Add data label to chart and set its id.
        let label1 = chart.Series.get_Item(0).DataLabels.Add();
        label1.ID = 0;

        // Set the default position of data label. This position is relative to the data markers.
        //label1.Position = ChartDataLabelPosition.OutsideEnd;

        // Set custom position of data label. This position is relative to the default position.
        label1.X = 0.1;
        label1.Y = -0.1;

        // Set label value visible
        label1.LabelValueVisible = true;

        // Set legend key invisible
        label1.LegendKeyVisible = false;

        // Set category name invisible
        label1.CategoryNameVisible = false;

        // Set series name invisible
        label1.SeriesNameVisible = false;

        // Set Percentage invisible
        label1.PercentageVisible = false;

        // Set border style and fill style of data label
        label1.Line.FillType = wasmModule.FillFormatType.Solid;
        label1.Line.SolidFillColor.Color = wasmModule.Color.get_Blue();
        label1.Fill.FillType = wasmModule.FillFormatType.Solid;
        label1.Fill.SolidColor.Color = wasmModule.Color.get_Orange();

        // Define the output file name
        const outputFileName = "SetPositionOfChartDataLabels.pptx";

        // Save the document 
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
