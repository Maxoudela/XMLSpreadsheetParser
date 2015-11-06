/*The MIT License (MIT)

 Copyright (c) 2015 Samir Hadzic

 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:

 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.

 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE.
 */
package com.github.maxoudela.xmlspreadsheetparser;

import java.nio.ByteBuffer;
import java.util.Set;
import javafx.application.Application;
import static javafx.application.Application.launch;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.input.Clipboard;
import javafx.scene.input.DataFormat;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

/**
 *
 * @author samir.hadzic
 */
public class ExcelClipBoardExtractor extends Application {

    public static void main(String[] args) {
        launch(args);
    }
    private static final String XML_SPREADSHEET_IDENTIFIER = "XML Spreadsheet";

    private static DataFormat findXmlFormat(Set<DataFormat> formats) {
        for (DataFormat format : formats) {
            if (format.getIdentifiers().contains(XML_SPREADSHEET_IDENTIFIER)) {
                return format;
            }
        }
        return null;
    }

    @Override
    public void start(Stage stage) {
        Label label = new Label("Copy something from Excel and click on the button to see the XML Spreadsheet content:");
        label.setStyle("-fx-font-weight: bold;");
        Button button = new Button("Click me");
        TextArea textArea = new TextArea();
        button.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                Clipboard clipBoard = Clipboard.getSystemClipboard();

                DataFormat xmlFormat = findXmlFormat(clipBoard.getContentTypes());

                if (xmlFormat != null && clipBoard.hasContent(xmlFormat)) {
                    String content = new String(((ByteBuffer) clipBoard.getContent(xmlFormat)).array());
                    textArea.setText(content);
                } else {
                    textArea.setText("Did you copy from Excel before? Nothing found in the clipBoard..");
                }
            }
        });
        VBox vbox = new VBox(label, button, textArea);
        vbox.setSpacing(10);
        vbox.setPadding(new Insets(10));
        Scene scene = new Scene(vbox);
        stage.setTitle("Excel clipboard extractor");
        stage.setScene(scene);
        stage.show();

    }
}
