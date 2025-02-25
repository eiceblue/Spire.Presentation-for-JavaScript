<template>
  <span>The example demonstrates how to get animation effect information. </span>
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

        const inputFileName = "Animation.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);
        
        let stringBuilder = [];

        //Travel each slide
        for(let i = 0;i < ppt.Slides.Count;i++){
            let slide = ppt.Slides.get_Item(i);
            for (let j = 0;j < slide.Timeline.MainSequence.Count;j++){
                let effect = slide.Timeline.MainSequence.get_Item(j);
                //Get the animation effect type
                let animationEffectType = effect.AnimationEffectType;
                stringBuilder.push("animation effect type:" + animationEffectType);

                //Get the slide number where the animation is located
                let slideNumber = slide.SlideNumber;
                stringBuilder.push("slide number:" + slideNumber);

                //Get the shape name
                let shapeName = effect.ShapeTarget.Name;
                stringBuilder.push("shape name:" + shapeName + "\n");
            }
        }


        // Define the output file name
        const outputFileName = "GetAnimationEffectInfo_out.txt";

        wasmModule.FS.writeFile(outputFileName, stringBuilder.join("\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],  {type: "text/plain"});

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
