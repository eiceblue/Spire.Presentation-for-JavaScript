<template>
  <span>Click the following button to set row height and column width for table.</span>
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

        let inputFileName = "SetRowHeightColumnWidth.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        //Get the table
        let table = null;
        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
            let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
            if (shape instanceof wasmModule.ITable) {
                table = shape;
                //Set the height for the rows
                table.TableRows.get_Item(0).Height = 100;
                table.TableRows.get_Item(1).Height = 80;
                table.TableRows.get_Item(2).Height = 60;
                table.TableRows.get_Item(3).Height = 40;
                table.TableRows.get_Item(4).Height = 20;

                //Set the column width
                table.ColumnsList.get_Item(0).Width = 60;
                table.ColumnsList.get_Item(1).Width = 80;
                table.ColumnsList.get_Item(2).Width = 120;
                table.ColumnsList.get_Item(3).Width = 140;
                table.ColumnsList.get_Item(4).Width = 160;
            }
        }

        // Define the output file name
        const outputFileName = "SetRowHeightColumnWidth_out.pptx";

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
