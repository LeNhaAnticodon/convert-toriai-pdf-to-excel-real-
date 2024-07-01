module com.example.convert_toriai_pdf_to_excel {
    requires javafx.controls;
    requires javafx.fxml;
            
                            
    opens com.example.convert_toriai_pdf_to_excel to javafx.fxml;
    exports com.example.convert_toriai_pdf_to_excel;
}