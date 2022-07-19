module com.FML {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.pdfbox;
    requires org.apache.poi.poi;
    requires java.desktop;
    requires org.apache.poi.ooxml;

    opens com.FML to javafx.fxml;
    exports com.FML;
}