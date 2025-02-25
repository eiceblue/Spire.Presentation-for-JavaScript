<template>
  <span>Click the following button to set text direction in a PPT document.</span>
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

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Append a shape with text to the first slide
        let textboxShape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType: wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(250, 70, 350, 470)});
        textboxShape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();
        textboxShape.Fill.FillType = wasmModule.FillFormatType.Solid;
        textboxShape.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
        textboxShape.TextFrame.Text = "You Are Welcome Here";
        
        // Set the text direction to vertical
        textboxShape.TextFrame.VerticalTextType = wasmModule.VerticalTextType.Vertical;

        // Append another shape with text to the slide
        textboxShape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType: wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(350, 70, 450, 470)});
        textboxShape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Transparent();
        textboxShape.Fill.FillType = wasmModule.FillFormatType.Solid;
        textboxShape.Fill.SolidColor.Color = wasmModule.Color.get_LightGray();

        // Append some asian characters
        textboxShape.TextFrame.Text = "欢迎光临";

        // Set the VerticalTextType as EastAsianVertical to aviod rotating text 90 degrees
        textboxShape.TextFrame.VerticalTextType = wasmModule.VerticalTextType.EastAsianVertical;

        const outputFileName = "SetTextDirection.pptx";

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
