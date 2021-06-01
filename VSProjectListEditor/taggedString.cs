namespace VSProjectListEditor
{
    public class TaggedString<T>
    {
        private readonly string m_title;
        public T Tag { get; private set; }

        public TaggedString(string title, T tag)
        {
            m_title = title;
            Tag = tag;
        }

        public override string ToString()
        {
            return m_title;
        }
    }
}
