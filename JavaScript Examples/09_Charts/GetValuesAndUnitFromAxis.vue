<template>
  <span>Click the following button to get values and unit from axis.</span>
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
        let inputFileName = "ChartSample2.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        let sb = [];

        //Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get chart on the first slide
        let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Get unit from primary category axis
        let MajorUnit = Chart.PrimaryCategoryAxis.MajorUnit;
        let type = Chart.PrimaryCategoryAxis.MajorUnitScale;

        sb.push(MajorUnit + "\r\n");
        sb.push(type + "\r\n");


        //Get values from primary value axis
        let minValue = Chart.PrimaryValueAxis.MinValue;
        let maxValue = Chart.PrimaryValueAxis.MaxValue;

        sb.push(minValue + "\r\n");
        sb.push(maxValue + "\r\n");

        // Define the output file name
        const outputFileName = "GetValuesAndUnitFromAxis_out.txt";

        // Save the document to the specified path
        FS.writeFile(outputFileName, sb.join(""))

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],  { type: 'text/plain' });

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
