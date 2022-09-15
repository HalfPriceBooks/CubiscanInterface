using CubiscanInterface.DBHelpers;

namespace CubiscanInterface
{
    class Program
    {
        static void Main(string[] args)
        {
            HPBInterface hpb = new HPBInterface();
            hpb.RunIngramData();
            hpb.PickUpFile();
            hpb.Update_SCALE();
            hpb.CleanUp(); 
        }
       
    }
}
