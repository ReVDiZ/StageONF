package Calculs;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class CalculInvEnPleinParcelle {

    private final String cheminFichierXLS;

    private POIFSFileSystem fileIn;
    private HSSFWorkbook workbook;
    private HSSFSheet sheet;
    private HSSFRow row;

    public CalculInvEnPleinParcelle(String cheminFichierXLS) {

        this.cheminFichierXLS = cheminFichierXLS;
    }

    public int lancerCalcul() {

        //On va chercher le fichier excel a chaque fois qu'on clique sur calcul
        //Ca va éviter à l'utilisateur de réimporter le fichier a chaque fois qu'il voudra changer une valeur de départ
        try {
            //Initialisation
            fileIn = new POIFSFileSystem(new FileInputStream(this.cheminFichierXLS));
            workbook = new HSSFWorkbook(fileIn);
            sheet = workbook.getSheetAt(0);
            row = sheet.getRow(0);

            //On verifie d'abord que la surface est bien rentrée
            int errSurface = verifSurface();
            if(errSurface != 0)
                return errSurface;

            //On calcul d'abord Nb pour chaque type d'arbre
            for(int i = 1; i <= 19; i += 3)
                sheet.getRow(26).getCell(i).setCellValue(getTotalNb(i));

            //On calcul Nb/Ha pour chaque type d'arbre
            for(int i = 1; i <= 19; i += 3)
                sheet.getRow(29).getCell(i).setCellValue(round2Dec(getTotalNbHa(i)));

            //On calcul G pour chaque type d'arbre
            for(int i = 1; i <= 19; i += 3)
                sheet.getRow(27).getCell(i).setCellValue(round2Dec(getTotalG(i)));

            //On calcul le diametre moyen pour chaque type d'arbre
            for(int i = 1; i <= 19; i += 3)
                sheet.getRow(28).getCell(i).setCellValue(round2Dec(getTotalDiametreMoyen(i)));

            //On calcul G pour chaque type de bois (PB, BM, GB, TGB) et pour chaque type d'arbre
            for(int ligne = 34; ligne <= 37; ligne++)
                for(int colonne = 5; colonne <= 11; colonne++)
                    sheet.getRow(ligne).getCell(colonne).setCellValue(round2Dec(getTotalG_Bois_Arbre(ligne, colonne)));
            //Puis on fait la somme pour avoir G pour chaque type de bois (PB, BM, GB, TGB) tout type d'arbre réunis
            for(int ligne = 34; ligne <= 37; ligne++)
                sheet.getRow(ligne).getCell(3).setCellValue(sommeG_Bois(ligne));

            //On calcul G' pour chaque type de bois (PB, BM, GB, TGB) (comme G mais sans DIV hors prod)
            for(int ligne = 34; ligne <= 37; ligne++)
                sheet.getRow(ligne).getCell(4).setCellValue(sommeGPrim_Bois(ligne));

            //On fait le total de G, G' et des G pour chaque type d'arbre
            for(int colonne = 3; colonne <= 11; colonne++)
                sheet.getRow(38).getCell(colonne).setCellValue(sommeTotaleBois(colonne));

            //On calcul le capital
            sheet.getRow(3).getCell(9).setCellValue(getCapital());

            //On calcul la structure
            sheet.getRow(2).getCell(9).setCellValue(getStructure());

            //On calcul la composition
            sheet.getRow(1).getCell(9).setCellValue(getComposition());

            //On calcul les volumes pour chaque type d'arbre
            for(int ligne = 9; ligne <= 25; ligne++)
                for(int colonne = 3; colonne <= 21; colonne += 3)
                    sheet.getRow(ligne).getCell(colonne).setCellValue(round2Dec(getVolume(ligne, colonne - 2)));

            //On calcul Vol/Ha pour chaque type d'arbre
            for(int i = 1; i <= 19; i += 3)
                sheet.getRow(30).getCell(i).setCellValue(round2Dec(getTotalVolHa(i)));
            
            //

            //A FAIRE - On calcul le RECPREV
            //MMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            //MMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            
            //On sauvegarde dans le fichier
            FileOutputStream fileOut;
            fileOut = new FileOutputStream(cheminFichierXLS);
            workbook.write(fileOut);
            fileOut.close();
        } catch(FileNotFoundException e) {
        } catch(IOException e) {
        }
        
        return 0;
    }

    private double getTotalNb(int numColonne) {
        double res = 0;
        for(int indiceLigne = 9; indiceLigne <= 25; indiceLigne++)
            res = res + sheet.getRow(indiceLigne).getCell(numColonne).getNumericCellValue();
        return res;
    }

    private double getTotalNbHa(int numColonne) {
        return sheet.getRow(26).getCell(numColonne).getNumericCellValue() / sheet.getRow(2).getCell(4).getNumericCellValue();
    }

    private double getTotalG(int numColonne) {
        double res = 0;
        double rayon;
        for(int indiceLigne = 9; indiceLigne <= 25; indiceLigne++) {
            rayon = sheet.getRow(indiceLigne).getCell(0).getNumericCellValue() / 200; //Pour l'avoir en metre
            res = res + sheet.getRow(indiceLigne).getCell(numColonne).getNumericCellValue() * Math.PI * Math.pow(rayon, 2);
        }
        res = res / sheet.getRow(2).getCell(4).getNumericCellValue();
        return res;
    }

    private double getTotalG_Bois_Arbre(int ligne, int colonne) {
        switch(ligne) {
            case 34:
                return getG_Bois_Arbre(9, 10, 1 + 3 * (colonne - 5));
            case 35:
                return getG_Bois_Arbre(11, 14, 1 + 3 * (colonne - 5));
            case 36:
                return getG_Bois_Arbre(15, 18, 1 + 3 * (colonne - 5));
            case 37:
                return getG_Bois_Arbre(19, 25, 1 + 3 * (colonne - 5));
            default:
                return 0;
        }
    }

    private double getG_Bois_Arbre(int ligneDebut, int ligneFin, int colonne) {
        double res = 0;
        double rayon;
        for(int indiceLigne = ligneDebut; indiceLigne <= ligneFin; indiceLigne++) {
            rayon = sheet.getRow(indiceLigne).getCell(0).getNumericCellValue() / 200; //Pour l'avoir en metre a partir du diametre
            res = res + sheet.getRow(indiceLigne).getCell(colonne).getNumericCellValue() * Math.PI * Math.pow(rayon, 2);
        }
        res = res / sheet.getRow(2).getCell(4).getNumericCellValue();
        return res;
    }

    private double sommeG_Bois(int ligne) {
        double res = 0;
        for(int colonne = 5; colonne <= 11; colonne++)
            res = res + sheet.getRow(ligne).getCell(colonne).getNumericCellValue();
        return res;
    }

    private double sommeGPrim_Bois(int ligne) {
        return sheet.getRow(ligne).getCell(3).getNumericCellValue() - sheet.getRow(ligne).getCell(11).getNumericCellValue();
    }

    private double sommeTotaleBois(int colonne) {
        double res = 0;
        for(int ligne = 34; ligne <= 37; ligne++)
            res = res + sheet.getRow(ligne).getCell(colonne).getNumericCellValue();
        return res;
    }

    private double getTotalDiametreMoyen(int numColonne) {
        double res = 0;
        double nb = 0;
        for(int indiceLigne = 9; indiceLigne <= 25; indiceLigne++) {
            res = res + sheet.getRow(indiceLigne).getCell(numColonne).getNumericCellValue() * sheet.getRow(indiceLigne).getCell(0).getNumericCellValue();
            nb = nb + sheet.getRow(indiceLigne).getCell(numColonne).getNumericCellValue();
        }
        if(nb == 0)
            return 0;
        else
            return res / nb;
    }

    private double getCapital() {
        return getCapitalDecimal();
    }

    private double getCapitalDecimal() {
        double res;
        double GPrim = sheet.getRow(38).getCell(4).getNumericCellValue();
        if(GPrim <= 7)
            res = 10 + getCapitalUnite();
        else if(GPrim <= 12)
            res = 20 + getCapitalUnite();
        else if(GPrim <= 17)
            res = 30 + getCapitalUnite();
        else if(GPrim <= 22)
            res = 40 + getCapitalUnite();
        else if(GPrim <= 30)
            res = 50 + getCapitalUnite();
        else
            res = 60 + getCapitalUnite();

        return res;
    }

    private double getCapitalUnite() {
        double res;
        double G = sheet.getRow(38).getCell(4).getNumericCellValue();
        if(G <= 7)
            res = 1;
        else if(G <= 12)
            res = 2;
        else if(G <= 17)
            res = 3;
        else if(G <= 22)
            res = 4;
        else if(G <= 30)
            res = 5;
        else
            res = 6;

        return res;
    }

    private String getStructure() {
        double GPrim_PB = sheet.getRow(34).getCell(4).getNumericCellValue();
        double GPrim_BM = sheet.getRow(35).getCell(4).getNumericCellValue();
        double GPrim_GB = sheet.getRow(36).getCell(4).getNumericCellValue();
        double GPrim_TGB = sheet.getRow(37).getCell(4).getNumericCellValue();
        double GPrim = sheet.getRow(38).getCell(4).getNumericCellValue();

        String code = "";

        if(GPrim_GB + GPrim_TGB <= 0.2 * GPrim)
            if(GPrim_BM <= 0.3 * GPrim)
                code = "11";
            else if(GPrim_BM <= 0.5 * GPrim)
                code = "12";
            else if(GPrim_BM <= 0.7 * GPrim)
                code = "21";
            else
                code = "22";

        if(GPrim_GB + GPrim_TGB > 0.2 * GPrim && GPrim_GB + GPrim_TGB <= 0.45 * GPrim)
            if(GPrim_BM <= 0.2 * GPrim)
                code = "13";
            else if(GPrim_BM <= 0.35 * GPrim)
                code = "51";
            else if(GPrim_PB < 0.1 * GPrim)
                code = "23";
            else
                code = "52";

        if(GPrim_GB + GPrim_TGB > 0.45 * GPrim && GPrim_GB + GPrim_TGB <= 0.75 * GPrim)
            if(GPrim_BM <= 0.2 * GPrim)
                code = "31";
            else if(GPrim_PB >= 0.1 * GPrim)
                code = "53";
            else
                code = "32";

        if(GPrim_GB + GPrim_TGB > 0.75 * GPrim)
            if(GPrim_GB > GPrim_TGB)
                code = "33G";
            else
                code = "33T";

        return getStructureCodeDefinition(code) + " : " + code;
    }

    private String getStructureCodeDefinition(String code) {
        String def = "Non défini";

        switch(code) {
            case "11":
                def = "Peuplement à Petits Bois";
                break;
            case "12":
                def = "Peuplement à Petits Bois avec Bois Moyens";
                break;
            case "21":
                def = "Peuplement à Bois Moyens avec Petits Bois";
                break;
            case "22":
                def = "Peuplement à Bois Moyens";
                break;
            case "13":
                def = "Peuplement à Petits Bois avec Gros Bois";
                break;
            case "51":
                def = "Peuplement Irrégulier à Petits Bois";
                break;
            case "23":
                def = "Peuplement à Bois Moyens avec Gros Bois";
                break;
            case "52":
                def = "Peuplement Irrégulier à Bois Moyens";
                break;
            case "31":
                def = "Peuplement à Gros Bois avec Petits Bois";
                break;
            case "53":
                def = "Peuplement Irrégulier à Gros Bois";
                break;
            case "32":
                def = "Peuplement à Gros Bois avec Bois Moyens";
                break;
            case "33G":
                def = "Peuplement à Gros Bois";
                break;
            case "33T":
                def = "Peuplement à Très Gros Bois";
                break;
        }

        return def;
    }

    private String getComposition() {
        double G_CHE = sheet.getRow(27).getCell(1).getNumericCellValue();
        double G_HET = sheet.getRow(27).getCell(4).getNumericCellValue();
        double G_FP = sheet.getRow(27).getCell(7).getNumericCellValue();
        double G_FRE = sheet.getRow(27).getCell(10).getNumericCellValue();
        double G_DIV = sheet.getRow(27).getCell(13).getNumericCellValue();
        double G_RSX = sheet.getRow(27).getCell(16).getNumericCellValue();
        double GPrim = sheet.getRow(38).getCell(4).getNumericCellValue();

        int compteurP1 = 0;
        int compteurP2 = 0;
        int indEssenceDominante = 0;
        int indEssenceDominee = 0;

        String typeArbre[] = {"CHE", "HET", "F.P", "FRE", "DIV", "RSX"};
        double g_Arbre[] = {G_CHE, G_HET, G_FP, G_FRE, G_DIV, G_RSX};

        for(int i = 0; i <= 5; i++)
            if(compteurP1 == 0)
                if(g_Arbre[i] > 0.65 * GPrim) {
                    compteurP1 = 1;
                    indEssenceDominante = i;
                } else if(g_Arbre[i] > 0.35 * GPrim) {
                    compteurP2 = compteurP2 + 1;
                    if(compteurP2 == 1)
                        indEssenceDominante = i;
                    else if(compteurP2 == 2)
                        if(g_Arbre[i] > g_Arbre[indEssenceDominante]) {
                            indEssenceDominee = indEssenceDominante;
                            indEssenceDominante = i;
                        } else
                            indEssenceDominee = i;
                }

        if(compteurP1 == 1)
            return "P1-" + typeArbre[indEssenceDominante];
        else if(compteurP2 == 2)
            return "P2-" + typeArbre[indEssenceDominante] + "-" + typeArbre[indEssenceDominee];
        else if(compteurP2 == 1)
            return "PM-" + typeArbre[indEssenceDominante];
        else
            return "PTM";
    }

    private double getVolume(int ligne, int colonne) {
        double res;
        //Si c'est la colonne RSX
        if(colonne == 16)
            //Si c'est la ligne de diametre 20
            if(ligne == 10)
                res = sheet.getRow(ligne).getCell(colonne).getNumericCellValue() * getVol_AMG_Résineux_Diam_inf_20(sheet.getRow(ligne).getCell(0).getNumericCellValue() / 100, sheet.getRow(ligne).getCell(colonne + 1).getNumericCellValue());
            else
                res = sheet.getRow(ligne).getCell(colonne).getNumericCellValue() * getVol_AMG_Résineux_Diam_sup_25(sheet.getRow(ligne).getCell(0).getNumericCellValue() / 100, sheet.getRow(ligne).getCell(colonne + 1).getNumericCellValue());
        else
            res = sheet.getRow(ligne).getCell(colonne).getNumericCellValue() * getVol_AMG_Feuillu(sheet.getRow(ligne).getCell(0).getNumericCellValue() / 100, sheet.getRow(ligne).getCell(colonne + 1).getNumericCellValue());

        return res;
    }

    private double getVol_AMG_Résineux_Diam_inf_20(double diam, double hauteur) {
        return (0.5 * Math.pow(diam, 2) * hauteur);
    }

    private double getVol_AMG_Résineux_Diam_sup_25(double diam, double hauteur) {
        return (0.45 * Math.pow(diam, 2) * hauteur);
    }

    private double getVol_AMG_Feuillu(double diam, double hauteur) {
        return (-0.027604 + 0.003582 * hauteur + 1.391321 * Math.pow(diam, 2) + 0.44846 * Math.pow(diam, 2) * hauteur);
    }

    private double getTotalVolHa(int colonne) {
        double volTotal = 0;
        for(int indiceLigne = 9; indiceLigne <= 25; indiceLigne++)
            volTotal = volTotal + sheet.getRow(indiceLigne).getCell(colonne + 2).getNumericCellValue();

        return volTotal / sheet.getRow(2).getCell(4).getNumericCellValue();
    }

    private double round2Dec(double d) {
        return Math.round(d * 100.0) / 100.0;
    }

    private int verifSurface() {
        if(sheet.getRow(2).getCell(4).getCellType() == HSSFCell.CELL_TYPE_BLANK)
            return -1;
        else if(sheet.getRow(2).getCell(4).getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
            return 0;
        else
            return -2;
    }
}
