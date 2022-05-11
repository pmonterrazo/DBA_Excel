import java.util.Scanner;

public class MaiorNumero{

    public static void main(String[] args){

        int num1;
        int num2;

        Scanner sc = new Scanner(System.in);

        System.out.print("Digite um numero: ");
        num1 = sc.nextInt();

        System.out.print("Digite um numero: ");
        num2 = sc.nextInt();

        sc.close();

        if(num1 > num2){
            System.out.println("O numero maior foi: " + num1);
        }else{
            System.out.println("O numero maior foi: " + num2);
        }
    }
}