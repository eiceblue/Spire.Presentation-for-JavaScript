<template>
  <span>Click the following button to use axis and format axis. </span>
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
        let inputFileName = 'ChartAxis.pptx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the chart
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Add a secondary axis to display the value of Series 3
        chart.Series.get_Item(2).UseSecondAxis = true;

        //Set the grid line of secondary axis as invisible
        chart.SecondaryValueAxis.MajorGridTextLines.FillType = wasmModule.FillFormatType.None;

        //Set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
        chart.PrimaryValueAxis.IsAutoMax = false;
        chart.PrimaryValueAxis.IsAutoMin = false;
        chart.SecondaryValueAxis.IsAutoMax = false;
        chart.SecondaryValueAxis.IsAutoMax = false;

        chart.PrimaryValueAxis.MinValue = 0;
        chart.PrimaryValueAxis.MaxValue = 5.0;
        chart.SecondaryValueAxis.MinValue = 0;
        chart.SecondaryValueAxis.MaxValue = 1.0;

        //Set axis line format
        chart.PrimaryValueAxis.MinorGridLines.FillType = wasmModule.FillFormatType.Solid;
        chart.SecondaryValueAxis.MinorGridLines.FillType = wasmModule.FillFormatType.Solid;
        chart.PrimaryValueAxis.MinorGridLines.Width = 0.1;
        chart.SecondaryValueAxis.MinorGridLines.Width = 0.1;
        chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = wasmModule.Color.get_LightGray();
        chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = wasmModule.Color.get_LightGray();
        chart.PrimaryValueAxis.MinorGridLines.DashStyle = wasmModule.LineDashStyleType.Dash;
        chart.SecondaryValueAxis.MinorGridLines.DashStyle = wasmModule.LineDashStyleType.Dash;

        chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3;
        chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = wasmModule.Color.get_LightSkyBlue();
        chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3;
        chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = wasmModule.Color.get_LightSkyBlue();
        // Define the output file name
        const outputFileName = "ChartAxis_out.pptx";

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
