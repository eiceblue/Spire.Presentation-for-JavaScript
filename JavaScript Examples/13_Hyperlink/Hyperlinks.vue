<template>
  <span>Click the following button to insert hyperlinks into a PPT document.</span>
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

        let ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        let rect =  wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle, fileName:ImageFileName,rectangle: rect});

        //Add new shape to PPT document
        let rec = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 255, 120, (500 + ppt.SlideSize.Size.Width / 2 - 255), 400);
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rec});
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.Line.Width = 0;

        //Add some paragraphs with hyperlinks
        let para1 = wasmModule.TextParagraph.Create();
        let tr = wasmModule.TextRange.Create("E-iceblue");
        tr.Fill.FillType = wasmModule.FillFormatType.Solid;
        tr.Fill.SolidColor.Color = wasmModule.Color.get_Blue();
        para1.TextRanges.Append(tr);
        para1.Alignment = wasmModule.TextAlignmentType.Center;
        shape.TextFrame.Paragraphs._Append(para1);
        shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

        //Add some paragraphs with hyperlinks
        let para2 = wasmModule.TextParagraph.Create();
        let tr1 = wasmModule.TextRange.Create("Click to know more about Spire.Presentation.");
        tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
        para2.TextRanges.Append(tr1);
        shape.TextFrame.Paragraphs._Append(para2);
        shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

        let para3 = wasmModule.TextParagraph.Create();
        let tr2 = wasmModule.TextRange.Create("Click to visit E-iceblue Home page.");
        tr2.ClickAction.Address = "https://www.e-iceblue.com/";
        para3.TextRanges.Append(tr2);
        shape.TextFrame.Paragraphs._Append(para3);
        shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

        let para4 = wasmModule.TextParagraph.Create();
        let tr3 = wasmModule.TextRange.Create("Click to go to the forum to raise questions.");
        tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html";
        para4.TextRanges.Append(tr3);
        shape.TextFrame.Paragraphs._Append(para4);
        shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

        let para5 = wasmModule.TextParagraph.Create();
        let tr4 = wasmModule.TextRange.Create("Click to contact our sales team via email.");
        tr4.ClickAction.Address = "mailto:sales@e-iceblue.com";
        para5.TextRanges.Append(tr4);
        shape.TextFrame.Paragraphs._Append(para5);
        shape.TextFrame.Paragraphs._Append(wasmModule.TextParagraph.Create());

        let para6 = wasmModule.TextParagraph.Create();
        let tr5 = wasmModule.TextRange.Create("Click to contact our support team via email.");
        tr5.ClickAction.Address = "mailto:support@e-iceblue.com";
        para6.TextRanges.Append(tr5);
        shape.TextFrame.Paragraphs._Append(para6);

        for (let i = 0; i < shape.TextFrame.Paragraphs.Count; i++) {
            let para = shape.TextFrame.Paragraphs.get_Item(i);
            if(para.Text.length != 0){
                para.TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Lucida Sans Unicode");
                para.TextRanges.get_Item(0).FontHeight = 20;
            }
        }


        // Define the output file name
        const outputFileName = "Hyperlinks_out.pptx";

        // Save the document to the specified path
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
