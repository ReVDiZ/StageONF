package Fenetres;

import Calculs.CalculInvEnPleinParcelle;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JLabel;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

public class MenuInvEnPleinParcelle extends JFrame implements ActionListener {

    private final JButton importXLS;
    private final JButton calcul;
    private final JLabel nomFichierXLS;

    private String cheminFichierXLS;

    private CalculInvEnPleinParcelle calculInvEnPleinParcelle;

    public MenuInvEnPleinParcelle() {
        //Titre de la fenetre
        this.setTitle("Inventaire en Plein - Parcelle");
        //Taille de la fenetre
        this.setSize(600, 300);
        //Positionne la fenetre au centre de l'écran
        this.setLocationRelativeTo(null);
        //Impossibilité de redimensionner la fenetre
        setResizable(false);
        //Quitte le process lorsqu'on clique sur la croix
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //JPanel global
        JPanel globalPanel = new JPanel();
        //On utilise un borderLayout pour placer les composants dedans par la suite
        globalPanel.setLayout(new BorderLayout());
        //Couleur de fond du JPanel
        globalPanel.setBackground(new Color(0, 100, 0));
        //Ajouter une bordure invisible au JPanel
        globalPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        //On prévient notre JFrame que notre JPanel sera son content pane
        this.setContentPane(globalPanel);

        //JLabel titre
        JLabel labelTitre = new JLabel("Inventaire en Plein - Parcelle");
        //Changer la police et la taille du JLabel
        labelTitre.setFont(new Font("Serif", Font.BOLD, 20));
        //Mettre le JLabel au milieu
        labelTitre.setForeground(Color.white);
        //Ajouter le JLabel titre au JPanel global
        labelTitre.setHorizontalAlignment(JLabel.CENTER);
        globalPanel.add(labelTitre, BorderLayout.NORTH);

        //JPanel qui contient le JPanel du gridLayout et le JLabel du nom du fichier
        JPanel panelCentre = new JPanel();
        panelCentre.setLayout(new BorderLayout());
        panelCentre.setBackground(new Color(0, 100, 0));
        //JPanel pour le grid Layout
        JPanel gridLayoutPanel = new JPanel();
        gridLayoutPanel.setBackground(new Color(0, 100, 0));
        gridLayoutPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        panelCentre.add(gridLayoutPanel, BorderLayout.NORTH);
        //On fait un GridLayout pour les boutons
        GridLayout inventairesGridLayout = new GridLayout(1, 2, 10, 10);
        gridLayoutPanel.setLayout(inventairesGridLayout);

        //On crée et redimensionne les Buttons
        importXLS = new JButton("Importer le fichier Excel");
        importXLS.addActionListener(this);
        calcul = new JButton("Déterminer les notations");
        calcul.addActionListener(this);
        nomFichierXLS = new JLabel("Import : Vide");
        nomFichierXLS.setForeground(Color.white);
        nomFichierXLS.setBorder(new EmptyBorder(0, 10, 10, 0));

        //Ajoute les Buttons au GridLayout
        gridLayoutPanel.add(importXLS);
        gridLayoutPanel.add(calcul);
        panelCentre.add(nomFichierXLS, BorderLayout.CENTER);

        globalPanel.add(panelCentre, BorderLayout.CENTER);

        //Ajoute le logo de l'ONF
        JLabel labelLogo = new JLabel();
        ImageIcon logoONF = new ImageIcon("Images/logoONF.png");
        labelLogo.setIcon(logoONF);
        labelLogo.setHorizontalAlignment(JLabel.CENTER);
        globalPanel.add(labelLogo, BorderLayout.SOUTH);

        //Rend visible la fenetre
        this.setVisible(true);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        Object source = e.getSource();

        if(source == importXLS) {
            //Init boite de dialogue et filtre
            FileFilter filter = new FileNameExtensionFilter("Fichiers .xls", "xls");
            JFileChooser dialogue = new JFileChooser();
            dialogue.setDialogTitle("Choisir un fichier Excel (xls)");
            dialogue.addChoosableFileFilter(filter);
            dialogue.setFileFilter(filter);
            //Affichage
            if(dialogue.showOpenDialog(null) == 0)
                if(filter.accept(dialogue.getSelectedFile())) {
                    //Récupération du fichier sélectionné
                    sauvegarderCheminXLS(dialogue.getSelectedFile().getPath());
                    changerLabelNomFichierXLS(dialogue.getSelectedFile().getName());
                    //Importer les données
                    calculInvEnPleinParcelle = new CalculInvEnPleinParcelle(cheminFichierXLS);
                } else {
                    sauvegarderCheminXLS(null);
                    changerLabelNomFichierXLS("Fichier non accepté");
                }
            else {
                sauvegarderCheminXLS(null);
                changerLabelNomFichierXLS("Annulé");
            }
        } else if(source == calcul)
            //calcul tout ce qu'il faut :)
            if(cheminFichierXLS != null) {
                //calcul tout ce qu'il faut, retourne -int si erreur de surface
                int err = calculInvEnPleinParcelle.lancerCalcul();
                if(err == -1/*probleme de surface vide*/)
                    popup("Vérifiez que vous avez rempli la surface");
                else if(err == -2/*probleme de surface non entier*/)
                    popup("Vérifiez que la surface est bien un nombre");
                else
                    //Dire : Ok j'ai fini :)
                    popup("Le calcul est terminé !\nVeuillez rafraichir votre fichier excel pour voir les changements.");
            } else
                popup("Vous n'avez pas importer de fichier excel (.xls)");
    }

    private void sauvegarderCheminXLS(String chemin) {
        this.cheminFichierXLS = chemin;
    }

    private void changerLabelNomFichierXLS(String nom) {
        nomFichierXLS.setText("Import : " + nom);
    }

    public void popup(String message) {
        JOptionPane.showMessageDialog(this, message);
    }
}
