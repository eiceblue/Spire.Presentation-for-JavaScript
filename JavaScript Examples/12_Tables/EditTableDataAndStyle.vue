<template>
  <span>The following example demonstrates how to edit table data and style in PowerPoint document</span>
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

        // Create PPT document and load file
        let ppt =wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        //Store the data used in replacement in string [].
        let str = ["Germany", "Berlin", "Europe", "0152458", "20860000"];

        //Get the table in PowerPoint document.
        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
          let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
          if (shape instanceof wasmModule.ITable) {
            //Change the style of table.
            shape.StylePreset =wasmModule.TableStylePreset.LightStyle1Accent2;

            for (let i = 0; i < shape.ColumnsList.Count; i++) {
              //Replace the data in cell.
              shape.get_Item(i, 2).TextFrame.Text = str[i];

              //Set the highlightcolor.
              shape.get_Item(i, 2).TextFrame.TextRange.HighlightColor.Color =wasmModule.Color.get_BlueViolet();
            }
          }
        }

        // Define the output file name
        const outputFileName = "EditTableDataAndStyle.pptx";

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
