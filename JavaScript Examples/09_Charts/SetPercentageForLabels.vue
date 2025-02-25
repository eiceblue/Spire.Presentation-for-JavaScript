<template>
  <span>The following example demonstrates how to add percentage label for stacked column chart</span>
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
        const inputFileName = "ColumnStacked.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);
        
        // Get the chart on the first slide
        let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        let dataPontPercent = 0;

        for (let i = 0; i < Chart.Series.Count; i++) {
          let series = Chart.Series.get_Item(i);
          //Get the total number
          let total = GetTotal(series.Values);
          for (let j = 0; j < series.Values.Count; j++) {
            //Get the percent
            dataPontPercent = parseFloat(series.Values.get_Item(j).Text) / total * 100;
            //Add datalabels
            let label = series.DataLabels.Add();
            label.LabelValueVisible = true;
            //Set the percent text for the label
            label.TextFrame.Paragraphs.get_Item(0).Text = `${dataPontPercent.toFixed(2)} %`;
            label.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).FontHeight = 12;
          }
        }

        // Define the output file name
        const outputFileName = "SetPercentageForLabels.pptx";

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
    function GetTotal( ranges) {
    let total = 0;
    for (let i = 0; i < ranges.Count; i++) {
        total += parseFloat(ranges.get_Item(i).Text);
    }
    return total;
}
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
