

package javaapplication1;


public class JavaApplication1 {

   
    public static void main(String[] args) {
        
       String s="ara#&vind";
       
       int j=0;
       for(int i=s.length()-1;i>=0;i--)
       {
          if(Character.isAlphabetic(s.charAt(j)))
                    System.out.print(s.charAt(i));
          else
              System.out.print(s.charAt(j));
          j++;
       }
       System.out.println();
       char[] arr = new char[s.length()];
       for(int i=s.length()-1;i>=0;i--)
       {
           if(!Character.isAlphabetic(s.charAt(i)))
               arr[i]=s.charAt(i);
        }
        int pos=0;
        
        for(int k=s.length()-1;k>=0;k--)
        {
           if(Character.isAlphabetic(s.charAt(k)))
           {
               while(arr[pos]!= '\0')
                   pos++;
           
                   arr[pos++]=s.charAt(k);
           }        
       }
        System.out.println(arr);
       
    }
    
}
