<template>
  <span>Click the following button to get the linked slide of the shape.</span>
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

        let inputFileName = "linkedSlide.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load ppt file
        ppt.LoadFromFile(inputFileName);
        let strB = [];

        //Get the second slide
        let slide = ppt.Slides.get_Item(1);

        //Get the first shape of the second slide
        let shape = slide.Shapes.get_Item(0);

        //Get the linked slide index
        if (shape.Click.ActionType == wasmModule.HyperlinkActionType.GotoSlide) {
            let targetSlide = shape.Click.TargetSlide;
            strB.push("Linked slide number = " + targetSlide.SlideNumber);
        }

        
        // Define the output file name
        const outputFileName = "GetLinkedSlide_out.txt";
        //Save
        FS.writeFile(outputFileName, strB.join(""));

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
