



import java.awt.FileDialog;
import java.io.File;
import java.io.IOException;

import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JTextArea;
import javax.swing.JTextPane;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

public class EficienciaFinanceiraCEF extends javax.swing.JFrame {
	
	static EFBB EFBB = new EFBB();
	private static GetSetBB getSetBB = new GetSetBB();
	
	static String nome = "";
	static String path = "";
	
	
	String labels[] = { "01 - JANEIRO", "02 - FEVEREIRO", "03 - MARCO", "04 - ABRIL","05 - MAIO", "06 - JUNHO", "07 - JULHO", "08 - AGOSTO","09 - SETEMBRO", "10 - OUTUBRO","11 - NOVEMBRO", "12 - DEZEMBRO" };
	String labelsAno[] = { "2015","2016","2017", "2018", "2019", "2020","2021", "2022", "2023", "2024","2025", "2026","2027", "2028", "2029", "2030" };
	

    public EficienciaFinanceiraCEF() {
        initComponents();
        setTitle("EFICIENCIA FINANCEIRA BANDO DO BRAZIL");
        
        UIManager.LookAndFeelInfo[] inf = UIManager.getInstalledLookAndFeels();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">                          
    private void initComponents() {

		JComboBox comboBox1 = new JComboBox(labels);
        comboBox1.setMaximumRowCount(12);
        
		JComboBox comboBox2 = new JComboBox(labelsAno);
        comboBox1.setMaximumRowCount(20);    	
    	

        jbtnBotaoExcel = new javax.swing.JButton();
        jbtnBotaoPDF = new javax.swing.JButton();

        jScrollPane1 = new javax.swing.JScrollPane();
        jtaArea = new javax.swing.JTextPane();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        comboBox2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
            	jbtnAnoActionPerformed(evt);
            }
        });
        
        
        comboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
            	jbtnMesActionPerformed(evt);
            }
        });        
        

        jbtnBotaoExcel.setText("SELECIONE O EXCELL");
        jbtnBotaoExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbtnBotaoExcelActionPerformed(evt);
            }
        });

        
        jbtnBotaoPDF.setText("SELECIONE O PDF");
        jbtnBotaoPDF.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbtnBotaoPDFActionPerformed(evt);
            }
        });

        jtaArea.setEditable(false);
       // jtaArea.setColumns(20);
        jtaArea.setFont(new java.awt.Font("Verdana", 0, 10)); // NOI18N
        //jtaArea.setRows(5);
        jScrollPane1.setViewportView(jtaArea);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1)
                    .addGroup(layout.createSequentialGroup()
                        //.addComponent(jbtnBotao1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(comboBox1)
                        .addComponent(comboBox2)
                        .addComponent(jbtnBotaoExcel)
                        .addComponent(jbtnBotaoPDF)
                        .addGap(18, 18, 18)))//

                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                		.addComponent(comboBox1)
                		.addComponent(comboBox2)                		
                    .addComponent(jbtnBotaoPDF)
                    .addComponent(jbtnBotaoExcel))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 124, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>                        

    
    private void jbtnAnoActionPerformed(java.awt.event.ActionEvent evt) {                                           
		JComboBox source = (JComboBox) evt.getSource();
    String item = (String) source.getSelectedItem();
		Ano(item);

}      

    private void jbtnMesActionPerformed(java.awt.event.ActionEvent evt) {                                           
    		JComboBox source = (JComboBox) evt.getSource();
        String item = (String) source.getSelectedItem();
    		Mes(item);

    }      
                                       

    private void jbtnBotaoExcelActionPerformed(java.awt.event.ActionEvent evt) {                                           
    		buscarFileDialogExcel();

    }    
    
    private void jbtnBotaoPDFActionPerformed(java.awt.event.ActionEvent evt) {                                           
        buscarFileDialogPDF();
    }                                          


                                              

    public static void main(String args[]) {
        (new EficienciaFinanceiraCEF()).show();
	      jtaArea.setText(""
	    		  +"\n   SIGA O PROCEDIMENTO ABAIXO: \n"
	    		  +"\n   1 - SELECIONE O EXCEL \n"
	    		  +"   2 - SELECIONE O PDF \n"
	    		  +"\n   APOS SELECIONAR O PDF AGUARDE ATÉ A MENSAGEM FINAL \n"
	    );
    }

    // Variables declaration - do not modify                     
    private javax.swing.JScrollPane jScrollPane1;
//    private javax.swing.JButton jbtnBotao1;
    private javax.swing.JButton jbtnBotaoExcel;
    private javax.swing.JButton jbtnBotaoPDF;
//    private javax.swing.JComboBox jcbxLookAndFeel;
    private static JTextPane jtaArea;
    // End of variables declaration                   


    
    private void buscarFileDialogExcel() {
        try {
            
                FileDialog fd = new FileDialog(this, "Buscar Texto", FileDialog.LOAD);
                fd.setMultipleMode(false);
                fd.show();

                File arquivo = new File(fd.getDirectory() + fd.getFile());

                if (!arquivo.isFile()) {
                    return;
                }
                
                nome = arquivo.getAbsolutePath();// pegando nome excel
              
                jtaArea.setText(""
      	    		  +"\n  MUITO BEM !!: \n"
      	    		  +"\n  AGORA SELECIONE O PDF \n"
      	    		  +"\n  APÓS SELECIONAR O PDF AGUARDE A MENSAGEM DE SUCESSO. \n"
                	);
                
               // jtaArea.setText("" +"\n \n \n \n SELECIONE O PDF E AGUARDE A MENSAGEM FINALIZANDO");              


                
                
        } catch (Exception e) {
        }
    }
    
    
    private void buscarFileDialogPDF() {
        try {
            
                FileDialog fd = new FileDialog(this, "Buscar Texto", FileDialog.LOAD);
                fd.setMultipleMode(false);
                fd.show();

                File arquivo = new File(fd.getDirectory() + fd.getFile());

                if (!arquivo.isFile()) {
                    return;
                }
                
                
                	 
                	jtaArea.setContentType("text/html"); 
                	jtaArea.setText("<html></body><center><h3><font color=#a70104>PROCESSANDO, AGUARDE ...</font>.</h3><img src=\"http://passofundo.ifsul.edu.br/imagens/padrao/aguarde.gif\"><center></body></html>"); 
                	

                	path = arquivo.getParent();
                rodarArquivos(path, nome);
                

                

                
                
        } catch (Exception e) {
        }
    }
    
    
    private void Ano(String Ano) {
        try {
        	
        		System.out.println("ANO BASE COMBO  :>"+ Ano);
        		
             
        		EFBB.ComboAno = Ano.trim();	

        } catch (Exception e) {
        }
    }
    
    
    
    private void Mes(String Mes) {
        try {
        	
        		System.out.println("MES BASE COMBO  :>"+ Mes.substring(0,2));
        		
        		Mes = Mes.substring(0,2);
             
        		EFBB.ComboMes = Mes.trim();	

        } catch (Exception e) {
        }
    }
        
    
    
    
    public void rodarArquivos(final String pasta, final String pathExcel) {
		new Thread() {
			
			@Override
			public void run() {
				
        		System.out.println("pathExcel --> "+pathExcel);
        		System.out.println("pasta --> "+pasta);
        		
             EFBB.excelBB = pathExcel;
             EFBB.caminho = pasta+"/";
             
             try 
             {
				EFBB.init();
				
             } catch (IOException e) 
             {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
             

             
         	jtaArea.setContentType("text/html"); 
         	jtaArea.setText("<html></body><center><h2><br><br><font color=#0056ee>FINALIZADO COM SUCESSO!!</font></h2><center></body></html>"); 
         	
			}
		}.start();

	}
}
