import java.util.Random;

/**
 * ClassName: rodom
 * Package: PACKAGE_NAME
 * Description:
 *
 * @Author fgb
 * @Create 2023/12/10 19:29
 * @Version 1.0
 */
public class rodom {
    public static void main(String[] args) {

        int[] arr = new int[100];

        for (int i = 0; i < arr.length; i++) {
            int num = (int) (Math.random()*2 + 1);
            arr[i] = num;
            System.out.print(arr[i]);
        }


    }
}
