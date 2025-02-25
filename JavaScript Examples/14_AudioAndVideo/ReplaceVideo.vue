<template>
  <span>Click the following button to replace an existing video in a PPT file.</span>
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

        let inputFileName2 = "repleaceVido.mp4";
        await wasmModule.FetchFileToVFS(inputFileName2,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the PPT document from disk.
        ppt.LoadFromFile(inputFileName);

        let videos = ppt.Videos;

        //Traverse all the slides of PPT file
        for (let i = 0; i < ppt.Slides.Count; i++) {
            let slide = ppt.Slides.get_Item(i);
            //Traverse all the shapes of slides
            for (let j = 0; j < slide.Shapes.Count; j++) {
                let shape = slide.Shapes.get_Item(j);
                //If shape is IVideo
                if (shape instanceof wasmModule.IVideo)
                {
                    //Replace the video
                    let video = shape;
                    //Load the video document from disk.
                    let bts = wasmModule.Stream.CreateByFile(inputFileName2);
                    let videoData = videos.Append({stream:bts});
                    video.EmbeddedVideoData = videoData;
                }
            }
        }

        // Define the output file name
        const outputFileName = "ReplaceVideo_out.pptx";

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
