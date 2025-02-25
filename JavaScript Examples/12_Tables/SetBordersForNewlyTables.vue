<template>
  <span>Click the following button to set the border type and color for newly added tables.</span>
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

        //Set the table width and height for each table cell.
        let tableWidth = [100, 100, 100, 100, 100]
        let tableHeight = [20, 20];

        //Traverse all the border type of the table.
        for (let item in wasmModule.TableBorderType) {
            //Add a table to the presentation slide with the setting width and height
            let itable = ppt.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight);

            //Add some text to the table cell.
            itable.TableRows.get_Item(0).get_Item(0).TextFrame.Text = "Row";
            itable.TableRows.get_Item(1).get_Item(0).TextFrame.Text = "Column";

            //Set the border type, border width and the border color for the table.
            itable.SetTableBorder(wasmModule.TableBorderType[item], 1.5,wasmModule.Color.get_Red());
        }


        // Define the output file name
        const outputFileName = "SetBordersForNewlyTables_out.pptx";

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
