/**
 * This class allows the user to select an ODS file and to send the info contained into it
 * into a Microsoft Access file. The header (first line) of the ODS file must contain the same fields as 
 * the selected table.
 * 
 * Arguments : 1. the ODS file, 2. the MS Access file, 3. the table name
 * 
 * @author David LECONTE
 * @version 1.0-SNAPSHOT
**/

package com.odstoaccess;

import java.sql.*;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

public class ODSToAccess {
    public static int START = XMLStreamConstants.START_ELEMENT;
    public static int END = XMLStreamConstants.END_ELEMENT;
    public static int CHARS = XMLStreamConstants.CHARACTERS;

    private Path filePath;
    private Path DBPath;
    private String DBTable;
    private String query;

    private XMLStreamReader reader;
    private Connection connection;

    private int linesRead;
    private Map<Integer, Integer> columnsReadPerLine;
    private int linesInserted;
    private Map<Integer, String> header;

    /**
     * Class constructor, sets and checks given files and tries to connect to MS
     * Access
     * 
     * @param fileLocation
     * @param DBPath
     * @param DBTable
     * @throws IOException
     * @throws SQLException
     * @throws ClassNotFoundException
     */
    public ODSToAccess(String fileLocation, String DBLocation, String DBTable)
            throws IOException, SQLException, ClassNotFoundException {

        this.filePath = Paths.get(fileLocation);

        if (!Files.exists(this.filePath))
            throw new FileNotFoundException("Specified ODS file does not exist !");

        this.DBPath = Paths.get(DBLocation);

        if (!Files.exists(this.DBPath))
            throw new FileNotFoundException("Specified MS Access database does not exist !");

        this.DBTable = DBTable;

        this.setXMLReaderFromODS();

        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        this.connection = DriverManager.getConnection("jdbc:ucanaccess://" + this.DBPath.toAbsolutePath());

        this.linesRead = 0;
        this.columnsReadPerLine = new HashMap<Integer, Integer>();
        this.linesInserted = 0;
    }

    /**
     * Sets the XML file reader from the given ODS file
     * 
     * @throws IOException
     */
    private void setXMLReaderFromODS() throws IOException {
        if (!this.filePath.toString().endsWith("ods") && !this.filePath.toString().endsWith("zip")) {
            throw new IllegalArgumentException("Wrong file format given, must end in .ods");
        }

        try {
            ZipFile zipFile = new ZipFile(this.filePath.toFile());
            ZipEntry zipEntry = zipFile.getEntry("content.xml");

            this.reader = XMLInputFactory.newInstance().createXMLStreamReader(zipFile.getInputStream(zipEntry));
        } catch (Exception e) {
            throw new IOException("Couldn't extract the ODS file, must be broken");
        }
    }

    /**
     * Starts the reading of all the ODS lines to send them 1 by 1 to the Access
     * table
     */
    public void readAllLines() {

        // allLines = new HashMap<Integer, HashMap<Integer, String>>();

        try {

            while (this.reader.hasNext()) {
                if (this.reader.next() == START && this.reader.getLocalName() == "table-row") {

                    if (this.linesRead == 0) {
                        this.header = this.readLine(this.reader);
                        this.linesRead++;
                        this.query = this.getQuery();
                    }

                    else {
                        Map<Integer, String> line = this.readLine(this.reader);
                        this.linesRead++;
                        this.insertLineIntoDB(line);
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Couldn't read through the lines of the table : " + e.getMessage());
        }
    }

    /**
     * Reads every single line of the ODS file from XML file extracted from it
     * 
     * @param reader
     * @return Each line as a Map
     * @throws XMLStreamException
     */
    private Map<Integer, String> readLine(XMLStreamReader reader) throws XMLStreamException {
        Map<Integer, String> mapSingleLine = new HashMap<Integer, String>();
        int lineNumber = this.linesRead;

        this.columnsReadPerLine.put(lineNumber, 0);

        while (reader.hasNext()) {

            int element = reader.next();

            if (element == START && reader.getLocalName() == "table-cell") {

                boolean readCell = readInnerCell(reader, mapSingleLine, lineNumber, false);

                if (!readCell)
                    break;
            }

            else if (element == END && reader.getLocalName() == "table-row")
                break;

            /*
             * for(HashMap.Entry<Integer, String> entry : mapSingleLine.entrySet()) {
             * System.out.println(entry.getKey() + " : " + entry.getValue() + " (" +
             * this.linesRead + ")"); } System.out.println();
             */
        }

        return mapSingleLine;
    }

    /**
     * Reads each single cell of the given ODS file
     * 
     * @param reader
     * @param mapSingleLine
     * @param lineNumber
     * @param recursion
     * @return True if a value was inserted in the line map, false if it is the end
     *         of the line
     * @throws XMLStreamException
     */
    private boolean readInnerCell(XMLStreamReader reader, Map<Integer, String> mapSingleLine, int lineNumber,
            boolean recursion) throws XMLStreamException {
        /*
         * If the element is empty, how many columns does the XML element count for ?
         * (number-columns-repeated attribute) Default is 1
         */
        int emptyColumnsRead = 1;

        if (reader.getEventType() == START && reader.getAttributeValue(null, "number-columns-repeated") != null) {
            emptyColumnsRead = Integer.parseInt(reader.getAttributeValue(null, "number-columns-repeated"));
        }

        while (this.reader.hasNext()) {

            int next1 = reader.next();

            // System.out.println("[" + lineNumber + "] {" +
            // this.columnsReadPerLine.get(lineNumber) + "}");

            // p equals text element here
            if (next1 == START && reader.getLocalName() == "p") {
                String tableCellText = "";

                while (reader.hasNext()) {
                    int next2 = reader.next();

                    if (next2 == END)
                        break;

                    else if (next2 == CHARS)
                        tableCellText += reader.getText();

                    else if (next2 == START)
                        tableCellText += recursiveElementRead(reader);
                }

                this.columnsReadPerLine.put(lineNumber, this.columnsReadPerLine.get(lineNumber) + 1);
                mapSingleLine.put(this.columnsReadPerLine.get(lineNumber), tableCellText);

                // System.out.println("[" + lineNumber + "] {" +
                // this.columnsReadPerLine.get(lineNumber) + "} " + tableCellText);

                return true;
            }

            else if (next1 == START && reader.getLocalName() == "table-cell") {
                this.columnsReadPerLine.put(lineNumber, this.columnsReadPerLine.get(lineNumber) + emptyColumnsRead);

                // System.out.println("[" + lineNumber + "] {" +
                // this.columnsReadPerLine.get(lineNumber) + "} EMPTY ADDED " +
                // emptyColumnsRead);

                return readInnerCell(reader, mapSingleLine, lineNumber, true);
            }

            else if (next1 == END && reader.getLocalName() == "table-row") {
                // System.out.println("end of row");
                this.columnsReadPerLine.put(lineNumber, this.columnsReadPerLine.get(lineNumber) + emptyColumnsRead);
                return false;
            }
        }

        return false;
    }

    /**
     * Method used recursively when an element contains other XML elements inside it
     * 
     * @param reader
     * @return The current read element also containing all the read elements inside
     *         it
     * @throws XMLStreamException
     */
    private String recursiveElementRead(XMLStreamReader reader) throws XMLStreamException {
        String text = "";
        String elementName = reader.getLocalName();

        while (reader.hasNext()) {
            int next2 = reader.next();

            if (next2 == END && reader.getLocalName() == elementName)
                break;

            else if (next2 == CHARS)
                text += reader.getText();

            else if (next2 == START)
                text += recursiveElementRead(reader);
        }

        return text;
    }

    /**
     * Generates query from header (1st line) fields given in ODS file
     * 
     * @return The assembled SQL query
     */
    private String getQuery() {
        int size = this.header.size();

        List<String> DBFields = new ArrayList<String>();

        this.query = "INSERT INTO " + this.DBTable + "(";
        int index = 1;

        for (Map.Entry<Integer, String> headerColumn : this.header.entrySet()) {
            String field = headerColumn.getValue();
            DBFields.add(field);

            this.query += "[" + field + "]";

            if (index < size)
                this.query += ",";
            index++;
        }

        this.query += ") VALUES (";

        for (int i = 0; i < size; i++) {
            this.query += "?";
            if (i < size - 1)
                this.query += ",";
        }

        this.query += ");";

        // System.out.println(this.query);

        return this.query;
    }

    /**
     * Performs the insertion of each line retrieved from the ODS file into the
     * selected table in the Access file
     * 
     * @param line
     * @return True if insertion succeeded, false otherwise
     */
    private boolean insertLineIntoDB(Map<Integer, String> line) {
        try {
            PreparedStatement preparedStatement = this.connection.prepareStatement(this.query);

            for (int i = 1; i <= this.header.size(); i++) {
                String field = line.get(i);

                if (field == null)
                    preparedStatement.setNull(i, java.sql.Types.VARCHAR);

                else
                    preparedStatement.setString(i, field);
            }

            if (preparedStatement.executeUpdate() > 0) {
                this.linesInserted++;
                if ((linesInserted % 20) == 0)
                    System.out.println(this.linesInserted + " rows inserted successfully !");
                return true;
            }
        } catch (Exception e) {
            int failedLine = linesInserted + 1;
            System.out.println("Prepared query for line " + failedLine + " failed : " + e.getMessage());
        }

        return false;
    }
}
