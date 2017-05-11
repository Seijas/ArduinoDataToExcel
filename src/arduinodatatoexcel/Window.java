package arduinodatatoexcel;

import com.panamahitek.ArduinoException;
import com.panamahitek.PanamaHitek_Arduino;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.table.DefaultTableModel;
import jssc.SerialPortEvent;
import jssc.SerialPortEventListener;
import jssc.SerialPortException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @version 1.0
 * @author Seijas
 */
public class Window extends javax.swing.JFrame {

    private final DefaultTableModel table;
    private final DateFormat hourFormat;
    private Date date;
    private String time;
    
    private boolean connect;
    private boolean state;
    private boolean tableCount;
    
    private final int baudios = 9600;
    
    PanamaHitek_Arduino ino;
    SerialPortEventListener listener = new SerialPortEventListener() {
        @Override
        public void serialEvent(SerialPortEvent spe) {
            try {
                if (ino.isMessageAvailable()){
                    inoData(ino.printMessage());
                }
            } catch (SerialPortException | ArduinoException ex) {
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    };
    
    public Window() {
        initComponents();
        table = (DefaultTableModel) jTableData.getModel();
        hourFormat = new SimpleDateFormat("HH:mm:ss");
        
        connect = false;
        state = false;
        tableCount = false;
        
        this.jButtonExport.setEnabled(tableCount);
        this.jLabelState.setText("Arduino not connected");
        
        writeComboBox();
        stateButtoms();
    }

    private void writeComboBox(){
        
        this.jComboBox.removeAllItems();
        
        this.jComboBox.addItem("COM2");
        this.jComboBox.addItem("COM3");
        this.jComboBox.addItem("COM5");
        
    }
    
    private void inoData(String data){
        
        ArduinoData d = new ArduinoData(data);
        
        date = new Date();
        time = hourFormat.format(date);
        
        table.addRow(new Object[]{time, d.getTemp(), d.getHum(), d.getState()});
        
        if (!tableCount) {
            if ( table.getRowCount() > 0 ){
                tableCount = true;
                this.jButtonExport.setEnabled(tableCount);
            }
        }
    }
    
    private String getCom(){
        return this.jComboBox.getSelectedItem().toString();
    }
    
    private void writeExcel(String path) {
        
        HSSFWorkbook libro = new HSSFWorkbook();
        HSSFSheet hoja = libro.createSheet();
        HSSFRow fila = hoja.createRow(0);
        HSSFCell celda = fila.createCell(0);
        
        celda.setCellValue("Arduino Incubator");
        
        // Se colocan los encabezados
        fila = hoja.createRow(1);
        celda = fila.createCell(0);
        celda.setCellValue("Time");
        celda = fila.createCell(1);
        celda.setCellValue("Temperature");
        celda = fila.createCell(2);
        celda.setCellValue("Humidity");
        celda = fila.createCell(3);
        celda.setCellValue("Light State");

        for (int i = 0; i <= table.getRowCount() - 1; i++) {
            
            fila = hoja.createRow(i + 2);
            
            for (int j = 0; j <= 3; j++) {
                celda = fila.createCell(j);
                celda.setCellValue(jTableData.getValueAt(i, j).toString());
            }
            
        }
        
        try {
            FileOutputStream Fichero;
            Fichero = new FileOutputStream(path);
            libro.write(Fichero);
            Fichero.close();
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void stateButtoms(){
        
        this.jButtonStart.setEnabled(connect);
        this.jButtonStop.setEnabled(connect);
        
    }
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanelPrincipal = new javax.swing.JPanel();
        jScrollPane = new javax.swing.JScrollPane();
        jTableData = new javax.swing.JTable();
        jButtonStart = new javax.swing.JButton();
        jButtonStop = new javax.swing.JButton();
        jButtonExport = new javax.swing.JButton();
        jPanelSuperior = new javax.swing.JPanel();
        jComboBox = new javax.swing.JComboBox<>();
        jButtonConnect = new javax.swing.JButton();
        jLabelState = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jTableData.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Time", "Temperature", "Humidity", "Light State"
            }
        ));
        jScrollPane.setViewportView(jTableData);

        jButtonStart.setText("Start recording");
        jButtonStart.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonStartActionPerformed(evt);
            }
        });

        jButtonStop.setText("Stop recording");
        jButtonStop.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonStopActionPerformed(evt);
            }
        });

        jButtonExport.setText("Export data tu excel");
        jButtonExport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonExportActionPerformed(evt);
            }
        });

        jComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jButtonConnect.setText("Connect");
        jButtonConnect.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonConnectActionPerformed(evt);
            }
        });

        jLabelState.setText("jLabel1");

        javax.swing.GroupLayout jPanelSuperiorLayout = new javax.swing.GroupLayout(jPanelSuperior);
        jPanelSuperior.setLayout(jPanelSuperiorLayout);
        jPanelSuperiorLayout.setHorizontalGroup(
            jPanelSuperiorLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanelSuperiorLayout.createSequentialGroup()
                .addComponent(jButtonConnect)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jLabelState, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, 78, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
        jPanelSuperiorLayout.setVerticalGroup(
            jPanelSuperiorLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelSuperiorLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(jComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addComponent(jButtonConnect)
                .addComponent(jLabelState))
        );

        javax.swing.GroupLayout jPanelPrincipalLayout = new javax.swing.GroupLayout(jPanelPrincipal);
        jPanelPrincipal.setLayout(jPanelPrincipalLayout);
        jPanelPrincipalLayout.setHorizontalGroup(
            jPanelPrincipalLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanelPrincipalLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanelPrincipalLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jPanelSuperior, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jScrollPane, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanelPrincipalLayout.createSequentialGroup()
                        .addComponent(jButtonStart)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButtonStop)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 35, Short.MAX_VALUE)
                        .addComponent(jButtonExport)))
                .addContainerGap())
        );
        jPanelPrincipalLayout.setVerticalGroup(
            jPanelPrincipalLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelPrincipalLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanelSuperior, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane, javax.swing.GroupLayout.DEFAULT_SIZE, 260, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanelPrincipalLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButtonStart)
                    .addComponent(jButtonStop)
                    .addComponent(jButtonExport))
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanelPrincipal, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanelPrincipal, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButtonStartActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonStartActionPerformed
        
        if (!state) {
            
            try {
                ino.sendData("1");
            } catch (ArduinoException | SerialPortException ex) {
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
            
            state = !state;
        }
        
    }//GEN-LAST:event_jButtonStartActionPerformed

    private void jButtonStopActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonStopActionPerformed
        
        if (state) {
            
            try {
                ino.sendData("0");
            } catch (ArduinoException | SerialPortException ex) {
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
            
            state = !state;
        }
        
    }//GEN-LAST:event_jButtonStopActionPerformed

    private void jButtonExportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonExportActionPerformed
        
        JFileChooser fch = new JFileChooser();
        
        if (fch.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            writeExcel(fch.getSelectedFile().getAbsolutePath() + ".xls");
        }
        
    }//GEN-LAST:event_jButtonExportActionPerformed

    private void jButtonConnectActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonConnectActionPerformed
        
        if (!connect) {
            
            try {
                ino = new PanamaHitek_Arduino();

                ino.arduinoRXTX(this.getCom(), baudios, listener);

                connect = true;
                
                this.jLabelState.setText("Arduino " + this.getCom() + " conected");
            
            } catch (ArduinoException ex) {
                connect = false;
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        } else {
            
            try {
                ino.killArduinoConnection();
                
                connect = false;
                
                this.jLabelState.setText("Arduino disconected...");
            
            } catch (ArduinoException ex) {
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        }
        
        if (connect) {
            this.jButtonConnect.setText("Disconnect");
        } else {
            this.jButtonConnect.setText("Connect");
        }
        
        this.stateButtoms();
        
    }//GEN-LAST:event_jButtonConnectActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButtonConnect;
    private javax.swing.JButton jButtonExport;
    private javax.swing.JButton jButtonStart;
    private javax.swing.JButton jButtonStop;
    private javax.swing.JComboBox<String> jComboBox;
    private javax.swing.JLabel jLabelState;
    private javax.swing.JPanel jPanelPrincipal;
    private javax.swing.JPanel jPanelSuperior;
    private javax.swing.JScrollPane jScrollPane;
    private javax.swing.JTable jTableData;
    // End of variables declaration//GEN-END:variables
}
