<template>
  <span>Click the following button to set the layout of the slide in a PPT document.</span>
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

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Remove the first slide
        ppt.Slides.RemoveAt(0);

        //Append a slide and set the layout for slide
        let slide = ppt.Slides.Append({ template: wasmModule.SlideLayoutType.Title });

        //Add content for Title and Text
        let shape = slide.Shapes.get_Item(0);
        shape.TextFrame.Text = "Hello Wolrd! -> This is title";

        shape = slide.Shapes.get_Item(1);
        shape.TextFrame.Text = "E-iceblue Support Team -> This is content";

        const outputFileName = "SetSlideLayout.pptx";

        // Save to file
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
