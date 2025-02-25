<template>
  <span>Click the following button to manage the header and footer of the node master.</span>
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

        let inputFileName = "PPTHasHeader.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load presentation
        ppt.LoadFromFile(inputFileName);

        //Set the note Masters header and footer
        let noteMasterSlide = ppt.NotesMaster;
        if (noteMasterSlide !== null) {
            for (let i = 0; i < noteMasterSlide.Shapes.Count; i++) {
                let shape =  noteMasterSlide.Shapes.get_Item(i);
                if (shape.Placeholder !== null) {
                    if (shape.Placeholder.Type == wasmModule.PlaceholderType.Header) {
                        shape.TextFrame.Text = "change the header by Spire";
                    }
                    if (shape.Placeholder.Type == wasmModule.PlaceholderType.Footer) {
                        shape.TextFrame.Text = "change the footer by Spire";
                    }
                }
            }
        }

        // Define the output file name
        const outputFileName = "ManageNoteMasterHeaderFooter_out.pptx";

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
