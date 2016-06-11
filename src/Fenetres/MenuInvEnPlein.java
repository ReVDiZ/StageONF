package Fenetres;

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
import javax.swing.border.EmptyBorder;

public class MenuInvEnPlein extends JFrame implements ActionListener {

    private MenuInvEnPleinParcelle menuInvEnPleinParcelle;
    //private MenuInvEnPleinBilanForet menuInvEnPleinBilanForet;

    private final JButton invEnPleinParParcelleButton;
    private final JButton invEnPleinGlobalButton;

    public MenuInvEnPlein() {
        //Titre de la fenetre
        this.setTitle("Inventaire en Plein");
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
        JLabel labelTitre = new JLabel("Inventaire en Plein");
        //Changer la police et la taille du JLabel
        labelTitre.setFont(new Font("Serif", Font.BOLD, 20));
        //Mettre le JLabel au milieu
        labelTitre.setForeground(Color.white);
        //Ajouter le JLabel titre au JPanel global
        labelTitre.setHorizontalAlignment(JLabel.CENTER);
        globalPanel.add(labelTitre, BorderLayout.NORTH);

        //JPanel pour le grid Layout
        JPanel gridLayoutPanel = new JPanel();
        gridLayoutPanel.setBackground(new Color(0, 100, 0));
        gridLayoutPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        globalPanel.add(gridLayoutPanel, BorderLayout.CENTER);
        //On fait un GridLayout pour les boutons
        GridLayout inventairesGridLayout = new GridLayout(1, 2, 10, 10);
        gridLayoutPanel.setLayout(inventairesGridLayout);

        //On crée et redimensionne les Buttons
        invEnPleinParParcelleButton = new JButton("Parcelle");
        invEnPleinParParcelleButton.addActionListener(this);
        invEnPleinGlobalButton = new JButton("Bilan Forêt");
        invEnPleinGlobalButton.addActionListener(this);

        //Ajoute les Buttons au GridLayout
        gridLayoutPanel.add(invEnPleinParParcelleButton);
        gridLayoutPanel.add(invEnPleinGlobalButton);

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

        if (source == invEnPleinParParcelleButton) {
            this.dispose();
            menuInvEnPleinParcelle = new MenuInvEnPleinParcelle();
        } else if (source == invEnPleinGlobalButton) {
            System.out.println("Inventaire En Plein Global");
        }
    }
}
