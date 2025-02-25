<template>
  <span>Click the following button to add note in a PPT document.</span>
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

        let inputFileName = "AddNote.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        ppt.LoadFromFile(inputFileName);

        let slide = ppt.Slides.get_Item(0);

        //Add note slide
        let notesSlide = slide.AddNotesSlide();

        //Add paragraph in the notesSlide
        let paragraph = wasmModule.TextParagraph.Create();
        paragraph.Text = "Tips for making effective presentations:";
        notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);

        paragraph = wasmModule.TextParagraph.Create();
        paragraph.Text = "Use the slide master feature to create a consistent and simple design template.";
        notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);
        //Set the bullet type for the paragraph in notesSlide
        notesSlide.NotesTextFrame.Paragraphs.get_Item(1).BulletType = wasmModule.TextBulletType.Numbered;
        notesSlide.NotesTextFrame.Paragraphs.get_Item(1).BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;

        paragraph = wasmModule.TextParagraph.Create();
        paragraph.Text = "Simplify and limit the number of words on each screen.";
        notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);
        notesSlide.NotesTextFrame.Paragraphs.get_Item(2).BulletType = wasmModule.TextBulletType.Numbered;
        notesSlide.NotesTextFrame.Paragraphs.get_Item(2).BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;

        paragraph = wasmModule.TextParagraph.Create();
        paragraph.Text = "Use contrasting colors for text and background.";
        notesSlide.NotesTextFrame.Paragraphs._Append(paragraph);
        notesSlide.NotesTextFrame.Paragraphs.get_Item(3).BulletType = wasmModule.TextBulletType.Numbered;
        notesSlide.NotesTextFrame.Paragraphs.get_Item(3).BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;

        // Define the output file name
        const outputFileName = "AddNote_out.pptx";

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
