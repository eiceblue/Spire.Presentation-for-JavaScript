<template>
  <span>Click the following button to traverse through the cells of table.</span>
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

        let content = [];
        content.push("The data in cells of this PowerPoint file is: ");

        //Get the table.
        let table = null;
        for (let i = 0; i < ppt.Slides.get_Item(0).Shapes.Count; i++) {
            let shape = ppt.Slides.get_Item(0).Shapes.get_Item(i);
            if (shape instanceof wasmModule.ITable) {
                table = shape;
                //Traverse through the cells of table.
                for (let j = 0; j < table.TableRows.Count; j++) {
                    let row = table.TableRows.get_Item(j);
                    for (let k = 0; k < row.Count; k++) {
                        let cell = row.get_Item(k);
                        content.push(cell.TextFrame.Text);
                    }
                    content.push("\n");
                }
            }
        }

        // Define the output file name
        const outputFileName = "TraverseThroughCells_out.txt";

        //Save to file
        FS.writeFile(outputFileName, content.join("\r\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "text/plain" });

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
