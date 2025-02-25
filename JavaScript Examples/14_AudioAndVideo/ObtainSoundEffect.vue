<template>
  <span>Click the following button to obtain sound effect of animation in a PPT document.</span>
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

        let inputFileName = "Animation.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load file
        ppt.LoadFromFile(inputFileName);

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Get the audio in a time node
        let audio = slide.Timeline.MainSequence.get_Item(0).TimeNodeAudios[0];

        //Get the properties of the audio, such as sound name, volume or detect if it's mute
        let sb = [];
        sb.push("SoundName: " + audio.SoundName);
        sb.push("Volume: " + audio.Volume);
        sb.push("IsMute: " + audio.IsMute);

        // Define the output file name
        const outputFileName = "ObtainSoundEffect_out.txt";

        //Save the properties of the audio to Text file
        FS.writeFile(outputFileName, sb.join("\r\n"));

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
