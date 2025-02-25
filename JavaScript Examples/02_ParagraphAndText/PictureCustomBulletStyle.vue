<template>
  <span>Click the following button to add picture as custom bullet style in a PPT document.</span>
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
        let inputFileName = "Bullets.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileImageName = "icon.png";
        await wasmModule.FetchFileToVFS(inputFileImageName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        //Get the second shape on the first slide
        let shape = ppt.Slides.get_Item(0).Shapes.get_Item(1);

        //Traverse through the paragraphs in the shape
        for (let i = 0; i < shape.TextFrame.Paragraphs.Count; i++) {
          let paragraph = shape.TextFrame.Paragraphs.get_Item(i);
          //Set the bullet style of paragraph as picture
          paragraph.BulletType = wasmModule.TextBulletType.Picture;
          //Load a picture
          let stream = wasmModule.Stream.CreateByFile(inputFileImageName);
          //Add the picture as the bullet style of paragraph
          paragraph.BulletPicture.EmbedImage = ppt.Images.Append({ stream: stream });
        }

        const outputFileName = "PictureCustomBulletStyle.pptx";

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
