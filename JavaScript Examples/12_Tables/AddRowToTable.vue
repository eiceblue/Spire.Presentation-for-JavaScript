<template>
  <span>The following example demonstrates how to add a row to table in PowerPoint document</span>
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
        const inputFileName = "Template_Ppt_1.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PowerPoint document.
        let ppt =wasmModule.Presentation.Create();

        // Load the file from VFS
        ppt.LoadFromFile(inputFileName);

        // Get the table within the PowerPoint document.
        let table = ppt.Slides.get_Item(0).Shapes.get_Item(0);

        // Get the first row.
        let row = table.TableRows.get_Item(1);

        // Clone the row and add it to the end of table.
        table.TableRows.Append(row);
        let rowCount = table.TableRows.Count;

        // Get the last row.
        let lastRow = table.TableRows.get_Item(rowCount - 1);

        // Set new data of the first cell of last row.
        lastRow.get_Item(0).TextFrame.Text = " The first added cell";

        // Set new data of the second cell of last row.
        lastRow.get_Item(1).TextFrame.Text = " The second added cell";

        // Define the output file name
        const outputFileName = "AddRowToTable.pptx";

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
