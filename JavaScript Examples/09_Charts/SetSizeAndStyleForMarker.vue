<template>
  <span>The following example shows how to set size and style for data marker</span>
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
        const inputFileName = "SetSizeAndStyleForMarker.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PowerPoint document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        // Get ther first chart
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        for (let i = 0; i < chart.Series.get_Item(0).Values.Count; i++) {
          //Create a ChartDataPoint object and specify the index.
          let dataPoint = wasmModule.ChartDataPoint.Create(chart.Series.get_Item(0));
          dataPoint.Index = i;

          //Set the fill color of the data marker.
          dataPoint.MarkerFill.Fill.FillType = wasmModule.FillFormatType.Solid;
          dataPoint.MarkerFill.Fill.SolidColor.Color = wasmModule.Color.get_Yellow();

          //Set the line color of the data marker.
          dataPoint.MarkerFill.Line.FillType = wasmModule.FillFormatType.Solid;
          dataPoint.MarkerFill.Line.SolidFillColor.Color = wasmModule.Color.get_YellowGreen();

          //Set the size of the data marker.
          dataPoint.MarkerSize = 20;

          //Set the style of the data marker
          dataPoint.MarkerStyle = wasmModule.ChartMarkerType.Diamond;
          chart.Series.get_Item(0).DataPoints.Add(dataPoint);
        }

        // Define the output file name
        const outputFileName = "SetSizeAndStyleForMarker.pptx";

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
