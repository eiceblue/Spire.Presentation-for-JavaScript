<template>
  <span>Click the following button to create map chart in a PPT document. </span>
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

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Insert a Map chart to the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.Map, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 500, 500), init: false });
        chart.ChartData._get_Item(0, 1).Text = "series";

        //Define some data.
        let countries = ["China", "Russia", "France", "Mexico", "United States", "India", "Australia"];
        for (let i = 0; i < countries.length; i++) {
          chart.ChartData._get_Item(i + 1, 0).Text = countries[i];
        }
        let values = [32, 20, 23, 17, 18, 6, 11];
        for (let i = 0; i < values.length; i++) {
          chart.ChartData._get_Item(i + 1, 1).NumberValue = values[i];
        }
        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 1, 0, 1);
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, 7, 0);
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 1, 7, 1);


        // Define the output file name
        const outputFileName = "CreateMapChart_out.pptx";

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
