<template>
  <span>Click the following button to set better transitions for slide.</span>
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
        let inputFileName = "SetTransitions.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName); 

        // Set the first slide transition as circle
        ppt.Slides.get_Item(0).SlideShowTransition.Type = wasmModule.TransitionType.Circle;

        // Set the transition time of 3 seconds
        ppt.Slides.get_Item(0).SlideShowTransition.AdvanceOnClick = true;
        ppt.Slides.get_Item(0).SlideShowTransition.AdvanceAfterTime = 3000;

        //Set the second slide transition as comb and set the speed
        ppt.Slides.get_Item(1).SlideShowTransition.Type = wasmModule.TransitionType.Comb;
        ppt.Slides.get_Item(1).SlideShowTransition.Speed = wasmModule.TransitionSpeed.Slow;

        // Set the transition time of 5 seconds
        ppt.Slides.get_Item(1).SlideShowTransition.AdvanceOnClick = true;
        ppt.Slides.get_Item(1).SlideShowTransition.AdvanceAfterTime = 5000;

        // Set the third slide transition as zoom
        ppt.Slides.get_Item(2).SlideShowTransition.Type = wasmModule.TransitionType.Zoom;

        // Set the transition time of 7 seconds
        ppt.Slides.get_Item(2).SlideShowTransition.AdvanceOnClick = true;
        ppt.Slides.get_Item(2).SlideShowTransition.AdvanceAfterTime = 7000;

        const outputFileName = "BetterSlideTransitions.pptx";

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
