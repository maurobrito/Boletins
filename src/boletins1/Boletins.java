package boletins1;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Boletins {

	static Workbook workbook;
	static WritableWorkbook copy;
	static String nome = "C", turma = "D", classificacao = "A", rm = "B", notafinal = "K", notach = "E", notacn = "F", nome2 = "C";
    static String notamat = "G", notaport = "H", notatotal = "I", notaredacao = "J", line;
	static Formula formulaNome, formulaTurma, formulaNotaCH, formulaNotaCN, formulaNotaMat, formulaNotaPort, formulaNotaTotal, formulaNotaRed, formulaNotaFinal;
	static Formula mediaCH, mediaCN, mediaMat, mediaPort, mediaTotal, mediaRed, mediaFinal, formulaNome2;
	
	public static void main(String[] args) throws BiffException, IOException, WriteException {
		CriarBoletins();
		Mensagem();
	}
	
	public static void CriarBoletins() throws RowsExceededException, WriteException, BiffException, IOException{
		workbook = Workbook.getWorkbook(new File("Modelo Boletim.xls"));
		copy = Workbook.createWorkbook(new File("output.xls"), workbook);
		WritableSheet sheet = copy.getSheet(0);
		
		int num_alunos = sheet.getRows();
		
		for(int i=1; i < num_alunos; i++){
			//pega a sheet na posição i para usá-la
			sheet = copy.getSheet(i);
			//faz line receber a linha que vai ser usada do total
			line = Integer.toString(i+37);
			
			//ajeita as strings para que elas possam receber a numeração correta
			nome = nome.substring(0, 1);
            turma = turma.substring(0, 1);
            classificacao = classificacao.substring(0, 1);
            rm = rm.substring(0, 1);
            notafinal = notafinal.substring(0, 1);
            notach = notach.substring(0, 1);
            notacn = notacn.substring(0, 1);
            notaport = notaport.substring(0, 1);
            notamat = notamat.substring(0, 1);
            notatotal = notatotal.substring(0, 1);
            notaredacao = notaredacao.substring(0, 1);
            nome2 = nome2.substring(0, 1);
            
            //concatena as strings para que elas mostrem a formula correta
            classificacao = classificacao + line;
            rm = rm + line;
            nome = nome + line;
            turma = turma + line;
            notach = notach + line;
            notacn = notacn + line;
            notamat = notamat + line;
            notaport = notaport + line;
            notatotal = notatotal + line;
            notaredacao = notaredacao + line;
            notafinal = notafinal + line;
            nome2 = nome2 + line;
            
            //cria as formulas das medias dos 20 primeiros
            mediaCH = new Formula(0, 18, "SUM(E38:E57)/20");
    		mediaCN = new Formula(1, 18, "SUM(F38:F57)/20");
    		mediaMat = new Formula(2, 18, "SUM(G38:G57)/20");
    		mediaPort = new Formula(3, 18, "SUM(H38:H57)/20");
    		mediaTotal = new Formula(4, 18, "SUM(I38:I57)/20");
    		mediaRed = new Formula(5, 18, "SUM(J38:J57)/20");
    		mediaFinal = new Formula(6, 18, "SUM(K38:K57)/20");
            
    		//cria as formulas das informações do aluno i
            formulaNome = new Formula(1, 7, nome);
            formulaTurma = new Formula(1, 8, turma);
            formulaNotaCH = new Formula(0, 12, notach);
            formulaNotaCN = new Formula(1, 12, notacn);
            formulaNotaMat = new Formula(2, 12, notamat);
            formulaNotaPort = new Formula(3, 12, notaport);
            formulaNotaTotal = new Formula(4, 12, notatotal);
            formulaNotaRed = new Formula(5, 12, notaredacao);
            formulaNotaFinal = new Formula(6, 12, notafinal);
            formulaNome2 = new Formula(1, 30, nome2);
            
            //adiciona as formulas com as informações do aluno à sheet do excel
            sheet.addCell(formulaNome);
            sheet.addCell(formulaTurma);
            sheet.addCell(formulaNotaCH);
            sheet.addCell(formulaNotaCN);
            sheet.addCell(formulaNotaMat);
            sheet.addCell(formulaNotaPort);
            sheet.addCell(formulaNotaTotal);
            sheet.addCell(formulaNotaRed);
            sheet.addCell(formulaNotaFinal);
            sheet.addCell(formulaNome2);
            
            //adiciona as formulas com as notas dos 20 primeiros à sheet do excel
            sheet.addCell(mediaCH);
            sheet.addCell(mediaCN);
            sheet.addCell(mediaMat);
            sheet.addCell(mediaPort);
            sheet.addCell(mediaTotal);
            sheet.addCell(mediaRed);
            sheet.addCell(mediaFinal);
            
            
		}
		//escreve na workbook do excel e fecha os arquivos
		copy.write();
		copy.close();
		workbook.close();
	}
	
	public static void Mensagem(){
	    System.out.println("Caso dê erros, veja se lembrou de colocar o arquivo em excel com o nome certo de 'Modelo Boletim.xls'");
	    System.out.println(" e com o resultado dos alunos e com o numero de abas a mais igual ao numero de alunos.");
	}

}
