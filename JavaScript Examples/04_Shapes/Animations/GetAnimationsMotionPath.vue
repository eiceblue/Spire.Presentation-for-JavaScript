<template>
  <span>The example shows how to get motion path of animations in PowerPoint document.</span>
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

        const inputFileName = "GetAnimationsMotionPath.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        let slide = ppt.Slides.get_Item(0);

        //Get the first shape
        let shape = slide.Shapes.get_Item(0);
        
        //Create a StringBuilder to save the tracks
        let StringBuilder = [];
        let o = 1;
        //Traverse all animations
        for(let i = 0;i < shape.Slide.Timeline.MainSequence.Count;i++){
            let effect = shape.Slide.Timeline.MainSequence.get_Item(i);
            if (effect.ShapeTarget.Equals(shape)) {
                //Get MotionPath
                let animationMotion = effect.CommonBehaviorCollection.get_Item(0);
                let path = animationMotion.Path;
                for (let j = 0;j < path.Count;j++){
                    let motionCmdPath = path.get_Item(j);
                    let points = motionCmdPath.Points;
                    let type = motionCmdPath.CommandType;
                    if(points != null){
                        for (let k = 0;k < points.length;k++){
                            let point = points[k];
                            StringBuilder.push(o + "  MotionType: " + type + " -> X: " + point.X + ", Y: " + point.Y);
                        }
                        o++;
                    }
                }
            }
        }

        // Define the output file name
        const outputFileName = "GetAnimationsMotionPath_out.txt";

        wasmModule.FS.writeFile(outputFileName, StringBuilder.join("\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
