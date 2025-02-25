<template>
  <span>The following example shows how to identify if it is merged cell</span>
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
        const inputFileName = "MergedCellInTable.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);
        
        //Get the first slide
        let slide = ppt.Slides.get_Item(0);
        let str = [];
        let output = "";
        for (let i = 0; i < slide.Shapes.Count; i++) {
          let shape = slide.Shapes.get_Item(i);
          //Verify if it is table
          if (shape instanceof wasmModule.ITable) {
            let table = shape;
            for (let r = 0; r < table.TableRows.Count; r++) {
              for (let c = 0; c < table.ColumnsList.Count; c++) {
                // Get cell
                let currentCell = table.TableRows.get_Item(r).get_Item(c);
                //Identify if it is merged cell
                if (currentCell.RowSpan > 1 || currentCell.ColSpan > 1) {
                  output = `Cell ${r}:${c} is a part of merged cell with RowSpan=${currentCell.RowSpan} and ColSpan=${currentCell.ColSpan} starting from Cell ${currentCell.FirstRowIndex}:${currentCell.FirstColumnIndex}.`;
                  str.push(output);
                }
              }
            }
          }
        }

        // Define the output file name
        const outputFileName = "IdentifyMergedCells.txt";

        // Read the saved file and convert to a Blob object
        const modifiedFile = new Blob([str.toString()], { type: "application/txt" });

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
