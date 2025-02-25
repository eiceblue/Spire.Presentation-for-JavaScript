<template>
  <span>The sample demonstrates how to extract images from specific slide in a PPT document</span>
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
        const inputFileName = "Images.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        imageDownloads.value = []; 
        let i = 0;
        //Traverse all shapes in the second slide
        for (let j = 0; j < ppt.Slides.get_Item(1).Shapes.Count; j++) {
          let shape = ppt.Slides.get_Item(1).Shapes.get_Item(j);
          //It is the SlidePicture object
          if (shape instanceof wasmModule.SlidePicture) {
            //Save to image
            let ps = shape;
            let fileName = `SlidePic_${i}.png`;

            let image = ps.PictureFill.Picture.EmbedImage.Image;
            image.Save(fileName);
            const imageFileArray = wasmModule.FS.readFile(fileName);
            const imageBlob = new Blob([imageFileArray], { type: "image/png" });

            // Add each image URL to the array for download
            imageDownloads.value.push({
              name: fileName,
              url: URL.createObjectURL(imageBlob),
            });

            image.Dispose();
            i++;
          }
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
