<template>
  <span>Click the following button to remove tick marks of axis.</span>
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
        let inputFileName = "Template_Ppt_2.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a PowerPoint document.
        let ppt = wasmModule.Presentation.Create();

        //Load the file from disk.
        ppt.LoadFromFile(inputFileName);

        //Get the chart that need to be adjusted the number format and remove the tick marks.
        let chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Set percentage number format for the axis value of chart.
        chart.PrimaryValueAxis.NumberFormat = "0#\\%";

        //Remove the tick marks for value axis and category axis.
        chart.PrimaryValueAxis.MajorTickMark = wasmModule.TickMarkType.TickMarkNone;
        chart.PrimaryValueAxis.MinorTickMark = wasmModule.TickMarkType.TickMarkNone;
        chart.PrimaryCategoryAxis.MajorTickMark = wasmModule.TickMarkType.TickMarkNone;
        chart.PrimaryCategoryAxis.MinorTickMark = wasmModule.TickMarkType.TickMarkNone;

        // Define the output file name
        const outputFileName = "RemoveTickMarksOfAxis_out.pptx";

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
