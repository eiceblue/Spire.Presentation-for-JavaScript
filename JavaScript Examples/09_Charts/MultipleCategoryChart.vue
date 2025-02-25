<template>
  <span>Click the following button to create multi-category chart in a PPT document.</span>
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

        //Create a PPT file
        let ppt = wasmModule.Presentation.Create();

        //Add line markers chart
        let rect1 = wasmModule.RectangleF.FromLTRB(90, 100, 640, 420);
        let chart = ppt.Slides.get_Item(0).Shapes.AppendChart({ type: wasmModule.ChartType.ColumnClustered, rectangle: rect1, init: false });

        //Chart title
        chart.ChartTitle.TextProperties.Text = "Muli-Category";
        chart.ChartTitle.TextProperties.IsCentered = true;
        chart.ChartTitle.Height = 30;
        chart.HasTitle = true;

        //Data for series
        let Series1 = [7.7, 8.9, 7, 6, 7, 8];

        //Set series text
        chart.ChartData._get_Item(0, 2).Text = "Series1";

        //Set category text
        chart.ChartData._get_Item(1, 0).Text = "Grp 1";
        chart.ChartData._get_Item(3, 0).Text = "Grp 2";
        chart.ChartData._get_Item(5, 0).Text = "Grp 3";

        chart.ChartData._get_Item(1, 1).Text = "A";
        chart.ChartData._get_Item(2, 1).Text = "B";
        chart.ChartData._get_Item(3, 1).Text = "C";
        chart.ChartData._get_Item(4, 1).Text = "D";
        chart.ChartData._get_Item(5, 1).Text = "E";
        chart.ChartData._get_Item(6, 1).Text = "F";


        //Fill data for chart
        for (let i = 0; i < Series1.length; ++i) {
          chart.ChartData._get_Item(i + 1, 2).NumberValue = Series1[i];
        }

        //Set series label
        chart.Series.SeriesLabel = chart.ChartData._get_ItemNE("C1", "C1");
        //Set category label
        chart.Categories.CategoryLabels = chart.ChartData._get_ItemNE("A2", "B7");

        //Set values for series
        chart.Series.get_Item(0).Values = chart.ChartData._get_ItemNE("C2", "C7");

        //Set if the category axis has multiple levels
        chart.PrimaryCategoryAxis.HasMultiLvlLbl = true;
        //Merge same label
        chart.PrimaryCategoryAxis.IsMergeSameLabel = true;


        // Define the output file name
        const outputFileName = "MultipleCategoryChart_out.pptx";

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
