namespace EpicPlanner;

internal class Wish
{
    #region Constructor

    public Wish(string _strResource, double _dPercentage)
    {
        Resource = _strResource;
        Percentage = _dPercentage;
    }

    #endregion

    #region Description

    public string Resource { get; }
    public double Percentage { get; }

    #endregion
}