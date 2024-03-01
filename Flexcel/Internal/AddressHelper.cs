namespace Flexcel.Internal;

public static class AddressHelper
{
    /// <summary>
    /// Получает адрес Excel ячейки (0,0) -> A1
    /// </summary>
    /// <param name="column">Номер колонки от 0</param>
    /// <param name="row">Номер строки от 0</param>
    /// <returns></returns>
    public static string GetCellAddress(int column, long row)
    {
        
        if (column! < 0)
            throw new ArgumentOutOfRangeException(nameof(column), "Should be more or equal zero");
        
        if (row < 0)
            throw new ArgumentOutOfRangeException(nameof(row), "Should be more or equal zero");
                
        const int columnsRoot = 'Z' - 'A' + 1;
        const int letterOffset = 'A';
        var rooted = GetRooted(column + 1, columnsRoot);
        var columnAddress = new string(rooted.Select(d => (char)(d + letterOffset)).Reverse().ToArray());
        return $"{columnAddress}{row+1}";
    }

    private static IEnumerable<int> GetRooted(int number, int root)
    {
        while (number > 0)
        {
            var remainder = (number - 1) % root;
            yield return remainder;
            number = (number - 1) / root;
        }
    }
}