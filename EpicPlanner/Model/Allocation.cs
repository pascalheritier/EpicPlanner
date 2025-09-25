namespace EpicPlanner;

internal class Allocation
{
    #region Constructor

    public Allocation(string _strEpic, int _iSprint, string _strResource, double _dHours, DateTime _Start)
    {
        Epic = _strEpic;
        Sprint = _iSprint;
        Resource = _strResource;
        Hours = _dHours;
        SprintStart = _Start;
    }

    #endregion

    #region Description

    public string Epic { get; }
    public int Sprint { get; }
    public string Resource { get; }
    public double Hours { get; }
    public DateTime SprintStart { get; }

    #endregion
}