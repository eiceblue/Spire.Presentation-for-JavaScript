<template>
  <span>Click the following button to copy a paragraph from one PPT document to another.</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName1 = "TextTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = "TextTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load the source file
        let ppt1 = wasmModule.Presentation.Create();
        ppt1.LoadFromFile(inputFileName1);
        
        // Get the text from the first shape on the first slide
        let sourceshp = ppt1.Slides.get_Item(0).Shapes.get_Item(0);
        const text = sourceshp.TextFrame.Text;

        // Load the target file
        let ppt2 = wasmModule.Presentation.Create();
        ppt2.LoadFromFile(inputFileName2);

        // Get the first shape on the first slide from the target file
        let destshp = ppt2.Slides.get_Item(0).Shapes.get_Item(0);

        // Add the text to the target file
        destshp.TextFrame.Text += "\n\n" + text;

        const outputFileName = "CopyParagraphToAnotherPPT.pptx";

        // Save to file
        ppt2.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

        // Clean up resources
        ppt2.Dispose();

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
