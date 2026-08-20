/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package jamie2026;

import java.util.Scanner;



/**
 *
 * @author Jamie Eduardo Rosal
 */
public class Jamie2026 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);

        int a = input.nextInt();
        switch (a) {
            case 1:
                System.out.println("welcome to case 1");
                System.out.println("Enter your number");
                int x = input.nextInt();
                System.out.println(Math.pow(x, 2) + " " + "the power");
                break;

            case 5:
                System.out.println("welcome to case 5");
                System.out.println("Enter your number");
                int y = input.nextInt();
                System.out.println(Math.pow(y, 3) + " " + "the power");
                break;
                
                
             case 7:
                int h=input.nextInt();
                 System.out.println("Enter a number");
                 if (h>=10) {
                     int f=input.nextInt();
                     if (f<100) {
                         System.out.println("less bai.. 100");
                     }else {
                         System.out.println("wow uli na..");
                     }
                 }
                
                break;

        }
    }

}
