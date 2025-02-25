<template>
  <span>The following example shows how to clone row and column of table</span>
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

        // Create PPT document 
        let ppt =wasmModule.Presentation.Create();

        // Access first slide
        let sld = ppt.Slides.get_Item(0);

        // Define columns with widths and rows with heights
        let widths = [110, 110, 110];
        let heights = [50, 30, 30, 30, 30];

        // Add table shape to slide
        let table = ppt.Slides.get_Item(0).Shapes.AppendTable(ppt.SlideSize.Size.Width / 2 - 275, 90, widths, heights);

        // Add text to the row 1 cell 1
        table.get_Item(0, 0).TextFrame.Text = "Row 1 Cell 1";

        // Add text to the row 1 cell 2
        table.get_Item(1, 0).TextFrame.Text = "Row 1 Cell 2";

        // Clone row 1 at end of table
        table.TableRows.Append(table.TableRows.get_Item(0));

        // Add text to the row 2 cell 1
        table.get_Item(0, 1).TextFrame.Text = "Row 2 Cell 1";

        // Add text to the row 2 cell 2
        table.get_Item(1, 1).TextFrame.Text = "Row 2 Cell 2";

        // Clone row 2 as the 4th row of table
        table.TableRows.Insert(3, table.TableRows.get_Item(1));

        //Clone column 1 at end of table
        table.ColumnsList.Add(table.ColumnsList.get_Item(0));

        //Clone the 2nd column at 4th column index
        table.ColumnsList.Insert(3, table.ColumnsList.get_Item(1));

        // Define the output file name
        const outputFileName = "CloneRowAndColumn.pptx";

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
