using ASCOM.Standard.Interfaces;

namespace Conform
{
    internal class SideOfPierResults
    {
        private PointingState m_SOP, m_DSOP;

        public SideOfPierResults()
        {
            m_SOP = PointingState.Unknown;
            m_DSOP = PointingState.Unknown;
        }

        internal PointingState SideOfPier
        {
            get
            {
                return m_SOP;
            }

            set
            {
                m_SOP = value;
            }
        }

        internal PointingState DestinationSideOfPier
        {
            get
            {
                return m_DSOP;
            }

            set
            {
                m_DSOP = value;
            }
        }
    }
}