<template>
  <span>The sample demonstrates how to extract images in a PPT document</span>
  <el-button @click="startProcessing">Start</el-button>
  <div v-if="imageDownloads.length">
    <h3>Click here to download the image:</h3>
    <ul>
      <li v-for="(image, index) in imageDownloads" :key="index">
        <a :href="image.url" :download="image.name">Download {{ image.name }}</a>
      </li>
    </ul>
  </div>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const imageDownloads = ref([]);

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);
         
        // Load the input file into the virtual file system (VFS)
        const inputFileName = "ExtractImage.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        imageDownloads.value = []; 

        for (let i = 0; i < ppt.Images.Count; i++) {
          let image = ppt.Images.get_Item(i).Image;
          let imageName = `Images_${i}.png`;

          // Save each image in virtual storage
          image.Save(imageName);
          const imageFileArray = wasmModule.FS.readFile(imageName);
          const imageBlob = new Blob([imageFileArray], { type: "image/png" });

          // Add each image URL to the array for download
          imageDownloads.value.push({
            name: imageName,
            url: URL.createObjectURL(imageBlob),
          });

          image.Dispose();
        }

        // Clean up resources
        ppt.Dispose();
      }
    };

    return {
      startProcessing,
      imageDownloads,
    };
  },
};
</script>
