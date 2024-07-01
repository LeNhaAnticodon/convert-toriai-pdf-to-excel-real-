module com.example.convert_toriai_from_pdf_to_excel {
    requires javafx.controls;
    requires javafx.fxml;
    requires com.opencsv;
    requires org.apache.pdfbox;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.example.convert_toriai_pdf_to_excel to javafx.fxml;
    exports com.example.convert_toriai_pdf_to_excel;
}