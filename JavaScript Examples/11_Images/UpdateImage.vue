<template>
  <span>The following example demonstrates how to replace image with new image in a PPT document</span>
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
         
        // Load the input file and image into the virtual file system (VFS)
        const inputFileName = "UpdateImage.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);
        const fileStreamName = "PresentationIcon.png";
        await wasmModule.FetchFileToVFS(fileStreamName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        // Append a new image to replace an existing image
        let stream = wasmModule.Stream.CreateByFile(fileStreamName);
        let image = ppt.Images.Append({ stream: stream });
        stream.Close();

        // Replace the image which title is "image1" with the new image
        for (let i = 0; i < slide.Shapes.Count; i++) {
          let shape = slide.Shapes.get_Item(i);
          if (shape instanceof wasmModule.SlidePicture) {
            if (shape.AlternativeTitle == "image1") {
              shape.PictureFill.Picture.EmbedImage = image;
            }
          }
        }

        // Define the output file name
        const outputFileName = "UpdateImage.pptx";

        // Save the document 
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
