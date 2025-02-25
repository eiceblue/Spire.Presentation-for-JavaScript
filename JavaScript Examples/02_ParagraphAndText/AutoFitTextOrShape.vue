<template>
  <span>Click the following button to auto fit text or shape in a PPT document.</span>
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

        // Load the image into the virtual file system (VFS)
        let imageName = "bg.png";
        await wasmModule.FetchFileToVFS(imageName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Create a rectangle from the left, top, right, and bottom coordinates based on slide size
        let rec = wasmModule.RectangleF.FromLTRB(0,0,ppt.SlideSize.Size.Width,ppt.SlideSize.Size.Height);
        
        // Append an embedded image as a rectangle shape to the first slide
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName:imageName,rectangle:rec});
        
        // Set the line color of the first shape to Floral White
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        // Set the AutofitType property to Shape
        let textShape2 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 300, 180)});
        
        // Add text in the shape
        textShape2.TextFrame.Text = "Resize shape to fit text.";
        textShape2.TextFrame.AutofitType = wasmModule.TextAutofitType.Shape;

        // Set the AutofitType property to Normal
        let textShape1 = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle:wasmModule.RectangleF.FromLTRB(400, 100, 550, 180)});
        textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape.";
        textShape1.TextFrame.AutofitType = wasmModule.TextAutofitType.Normal;

        // Define the output file name 
        const outputFileName = "AutoFitTextOrShape.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

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
