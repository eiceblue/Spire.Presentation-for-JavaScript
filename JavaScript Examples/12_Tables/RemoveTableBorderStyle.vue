<template>
  <span>Click the following button to remove table border style in PowerPoint document.</span>
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

        let inputFileName = "Template_Ppt_1.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the file from disk.
        ppt.LoadFromFile(inputFileName);

        for (let o = 0; o < ppt.Slides.Count; o++) {
            let slide = ppt.Slides.get_Item(o);
            for (let i = 0; i < slide.Shapes.Count; i++) {
                let shape = slide.Shapes.get_Item(i);
                //Verify if it is table
                if (shape instanceof wasmModule.ITable) {
                    let table = shape;
                    for (let j = 0; j < table.TableRows.Count; j++) {
                        let row = table.TableRows.get_Item(j);
                        for (let k = 0; k < row.Count; k++) {
                            let cell = row.get_Item(k);
                            cell.BorderTop.FillType = wasmModule.FillFormatType.None;
                            cell.BorderBottom.FillType = wasmModule.FillFormatType.None;
                            cell.BorderLeft.FillType = wasmModule.FillFormatType.None;
                            cell.BorderRight.FillType = wasmModule.FillFormatType.None;
                        }
                    }
                }
            }
        }

        // Define the output file name
        const outputFileName = "RemoveTableBorderStyle_out.pptx";

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
