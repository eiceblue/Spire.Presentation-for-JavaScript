<template>
  <span>Click the following button to create Treemap chart. </span>
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

        //Create PPT document
        let ppt = wasmModule.Presentation.Create();

        //Create a TreeMap chart to the first slide
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.TreeMap, rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 550, 450), init: false });

        //Set series text
        chart.ChartData._get_Item(0, 3).Text = "Series 1";

        //Set category text
        let categories = [["Branch 1", "Stem 1", "Leaf 1"], ["Branch 1", "Stem 1", "Leaf 2"], ["Branch 1", "Stem 1", "Leaf 3"],
        ["Branch 1", "Stem 2", "Leaf 4"], ["Branch 1", "Stem 2", "Leaf 5"], ["Branch 1", "Stem 2", "Leaf 6"], ["Branch 1", "Stem 2", "Leaf 7"],
        ["Branch 2", "Stem 3", "Leaf 8"], ["Branch 2", "Stem 3", "Leaf 9"], ["Branch 2", "Stem 4", "Leaf 10"], ["Branch 2", "Stem 4", "Leaf 11"],
        ["Branch 2", "Stem 5", "Leaf 12"], ["Branch 3", "Stem 5", "Leaf 13"], ["Branch 3", "Stem 6", "Leaf 14"], ["Branch 3", "Stem 6", "Leaf 15"]];
        for (let i = 0; i < categories.length; i++) {
          for (let j = 0; j < categories[0].length; j++) {
            chart.ChartData._get_Item(i + 1, j).Text = categories[i][j];
          }
        }

        //Fill data for chart
        let values = [17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51];
        for (let i = 0; i < values.length; i++) {
          chart.ChartData._get_Item(i + 1, 3).NumberValue = values[i];
        }

        //Set series labels
        chart.Series.SeriesLabel = chart.ChartData._get_ItemRCLL(0, 3, 0, 3);

        //Set categories labels
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemRCLL(1, 0, values.length, 2);

        //Assign data to series values
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemRCLL(1, 3, values.length, 3);

        chart.Series.get_Item(0).DataLabels.CategoryNameVisible = true;
        chart.Series.get_Item(0).TreeMapLabelOption = wasmModule.TreeMapLabelOption.Banner;
        chart.ChartTitle.TextProperties.Text = "TreeMap";
        chart.HasLegend = true;
        chart.ChartLegend.Position = wasmModule.ChartLegendPositionType.Top;

        // Define the output file name
        const outputFileName = "CreateTreeMapChart_out.pptx";

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
