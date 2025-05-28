package us.dit.fs.gestordocumental;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;


class TestPlantillaWord {
	    private static final Logger logger = LogManager.getLogger(TestPlantillaWord.class);
	    static PlantillaWord wordDocument;
	    
	    @BeforeAll
	    public static void generateMSWordFile() throws Exception {
	    	/**
	    	 * El test comienza creando un documento Word llamado Quijote.docx
	    	 * El título será "Capítulo primero"
	    	 * El texto del resto del documento depende del contenido de los ficheros subtitulo.txt y parrafo.txt
	    	 */
	        TestPlantillaWord.wordDocument = new PlantillaWord("Quijote.docx");
	        wordDocument.addTitle("Capítulo primero");
	        wordDocument.addSubtitle("subtitulo.txt");
	        wordDocument.addParagraph("parrafo.txt");
	        wordDocument.finishDocument();
	    }
	    /**
	     * Verifica el contenido del documento Word creado
	     * @throws Exception
	     */
	    @Test
	    public void whenParsingOutputDocument_thenCorrect() throws Exception {
	        Path msWordPath = Paths.get("Quijote.docx");
	        logger.info("path ",msWordPath);
	        XWPFDocument document = new XWPFDocument(Files.newInputStream(msWordPath));
	        List<XWPFParagraph> paragraphs = document.getParagraphs();
	        document.close();
	        assertEquals("Capítulo primero", paragraphs.get(0).getText());	        
	        assertEquals(wordDocument.convertTextFileToString("subtitulo.txt"), paragraphs.get(1).getText());
	        assertEquals(wordDocument.convertTextFileToString("parrafo.txt"), paragraphs.get(2).getText());
	       
	    }
		/**
		 * Verifica el tamaño del subtitulo, que debe ser 16
		 * @throws Exception
		 */
		@Test
		public void checkSubtitlesFontSize() throws Exception{
			Path msWordPath = Paths.get("Quijote.docx");
			logger.info("path ",msWordPath);
			XWPFDocument document = new XWPFDocument(Files.newInputStream(msWordPath));
			List<XWPFParagraph> paragraphs = document.getParagraphs();
			document.close();
			assertEquals(16, paragraphs.get(1).getRuns().get(0).getFontSize());

		}
		
		@Test
	    public void checkSubtitleFontSize() throws Exception {
	        Path path = Paths.get("Quijote.docx");
	        try (XWPFDocument document = new XWPFDocument(Files.newInputStream(path))) {
	            List<XWPFParagraph> paragraphs = document.getParagraphs();
	            int fontSize = paragraphs.get(1).getRuns().get(0).getFontSize();

	            int expectedFontSize = 16;
	            // Mensaje personalizado que muestra expected y actual
	            String message = String.format("Expected font size: %d, but was: %d", expectedFontSize, fontSize);
	            //System.out.println(message);

	            assertEquals(expectedFontSize, fontSize, message);
	        }
	    }
	}





   
