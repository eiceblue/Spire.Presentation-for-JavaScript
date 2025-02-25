<template>
  <span>Click the following button to extract videos from a PPT file.</span>
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

        let inputFileName = "video.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT document from disk.
        ppt.LoadFromFile(inputFileName);

        //Define a variable
        let i = 0;

        //String for output file
        let result = `ExtractVideo_${i}.avi`;
        //Traverse all the slides of PPT file
        for (let j = 0; j < ppt.Slides.Count; j++) {
            let slide = ppt.Slides.get_Item(j);
            //Traverse all the shapes of slides
            for (let k = 0; k < slide.Shapes.Count; k++) {
                let shape = slide.Shapes.get_Item(k);
                //If shape is IVideo
                if (shape instanceof wasmModule.IVideo)
                {
                    //Save the video
                    shape.EmbeddedVideoData.SaveToFile(result);
                    i++;
                }
            }
        }

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(result);
        const modifiedFile = new Blob([modifiedFileArray], { type: "video/x-msvideo" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = result;
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
