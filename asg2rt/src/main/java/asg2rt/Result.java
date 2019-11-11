package asg2rt;


public class Result
{
    public static void main( String[] args )
    {
        System.out.println("| Login id                     | Number of repositories | Number of followers | Number of stars  | Number of following   |");
        System.out.println("|------------------------------|------------------------|---------------------|------------------|-----------------------|");

        ConvertExcel save = new ConvertExcel();
        save.saveData();
        Stars acc1 = new Stars();
        acc1.showAcc1();
        Following acc2 = new Following();
        acc2.showAcc2();
        Repo acc3 = new Repo();
        acc3.showAcc3();

        System.out.println(" ");
        System.out.println("Storing data into excel");
        System.out.println("Successfully");
    }
}
