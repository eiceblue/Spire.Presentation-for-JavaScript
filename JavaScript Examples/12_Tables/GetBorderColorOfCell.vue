<template>
  <span>The following example shows how to get the disply color and border color of table cells in PowerPoint
    document</span>
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
        const inputFileName = "GetBorderColorOfCell.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);
        //Get the table in the first slide
        let table = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        //Get borders' color of the first cell
        let sb = [];
        sb.push("Color of left border:" + table.get_Item(0, 0).BorderLeftDisplayColor.ToString());
        sb.push("Color of top border:" + table.get_Item(0, 0).BorderTopDisplayColor.ToString());
        sb.push("Color of right border:" + table.get_Item(0, 0).BorderRightDisplayColor.ToString());
        sb.push("Color of bottom border:" + table.get_Item(0, 0).BorderBottomDisplayColor.ToString());

        //Get display color of the first cell
        sb.push("Color of cell:" + table.get_Item(0, 0).DisplayColor.ToString());

        // Define the output file name
        const outputFileName = "GetBorderColorOfCell.txt";

        // Read the saved file and convert to a Blob object
        const modifiedFile = new Blob([sb.toString()], { type: "application/txt" });

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
