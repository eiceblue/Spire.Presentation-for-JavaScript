<template>
  <span>The following example demonstrates how to convert a PPT document to SVG files</span>
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
        const inputFileName = "ToSVG.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        imageDownloads.value = [];
        for (let i = 0; i < ppt.Slides.Count; i++) {
          let svgBytes = ppt.Slides.get_Item(i).SaveToSVG();
          let fileName = `ToSVG-${i}.svg`;

          // Save each image in virtual storage
          svgBytes.Save(fileName);
          const imageFileArray = wasmModule.FS.readFile(fileName);
          const imageBlob = new Blob([imageFileArray], { type: "image/svg" });

          // Add each image URL to the array for download
          imageDownloads.value.push({
            name: fileName,
            url: URL.createObjectURL(imageBlob),
          });

          svgBytes.Dispose();
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
