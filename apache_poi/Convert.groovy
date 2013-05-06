
import org.apache.poi.poifs.filesystem.*
import org.apache.poi.hwpf.*
import org.apache.poi.hwpf.converter.*

import javax.xml.parsers.*
import javax.xml.transform.*
import javax.xml.transform.dom.*
import javax.xml.transform.stream.*

def buildWordToHtmlConverter = { def filename ->
    def fs = new POIFSFileSystem(new FileInputStream(filename))
    def hwpfDocument = new HWPFDocument(fs)
    def newDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument()
    def wordToHtmlConverter = new WordToHtmlConverter(newDocument)
    wordToHtmlConverter.processDocument( hwpfDocument )
    return wordToHtmlConverter    
}

def buildHtml = { def wordToHtmlConverter ->
    def stringWriter = new StringWriter()

    def transformer = TransformerFactory.newInstance().newTransformer()
    transformer.setOutputProperty( OutputKeys.INDENT, "yes" )
    transformer.setOutputProperty( OutputKeys.ENCODING, "utf-8" )
    transformer.setOutputProperty( OutputKeys.METHOD, "html" )
    transformer.transform( new DOMSource( wordToHtmlConverter.getDocument() ),
                           new StreamResult( stringWriter ) )

    return stringWriter.toString()
}

def usage = { ->
    println "usage: groovy Convert.groovy [filename]"
}

// ----------------------           main

if (args.length >= 1) {
    def filename = args[0]
    def wordToHtmlConverter = buildWordToHtmlConverter(filename)
    def html = buildHtml(wordToHtmlConverter)

    println html    
} else {
    usage()
}
