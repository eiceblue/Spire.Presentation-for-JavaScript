<template>
  <span>Click the following button to add and get speaker notes of chart.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";
import JSZip from "jszip";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName = "Template_Ppt_1.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let outputDirectoryName = "outputFiles/";
        FS.mkdirTree(outputDirectoryName);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load the file from disk.
        ppt.LoadFromFile(inputFileName);

        //Get the first slide and in the PowerPoint document.
        let slide = ppt.Slides.get_Item(0);

        //Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
        let ns = slide.NotesSlide;

        if (ns.H == undefined) {
            ns = slide.AddNotesSlide();
        }

        //Add the text string as the notes.
        ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation";

        let content = [];
        content.push("The speaker notes added by Spire.Presentation is: " + ns.NotesTextFrame.Text);

        let outputFile_txt = "GetSpeakerNotes.txt"
        //Get the speaker notes and save to txt file.
        FS.writeFile(outputDirectoryName+outputFile_txt, content.join(""));

        // Define the output file name
        const outputFileName = "AddAndGetSpeakerNotes_out.pptx";

        // Save the document to the specified path
        ppt.SaveToFile({ file: outputDirectoryName+outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

        // Clean up resources
        ppt.Dispose();

        const zip = new JSZip();
        let items = await FS.readdir(outputDirectoryName);
        items = items.filter((item) => item !== "." && item !== "..");
        for (const item of items) {
          const itemPath = `${outputDirectoryName}/${item}`;
          const fileData = await FS.readFile(itemPath);
          zip.file(item, fileData);
        }

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `outputFiles.zip`;
        downloadName.value = zipDownloadName;
        downloadUrl.value = zipDownloadUrl;


        // // Read the saved file and convert to a Blob object
        // const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        // const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });



        // // Download the file
        // downloadName.value = outputFileName;
        // downloadUrl.value = URL.createObjectURL(modifiedFile);
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
