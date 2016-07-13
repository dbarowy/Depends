using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Depends
{
    public class SparseMatrix
    {
        // represents a uniform-cost directed graph
        // first index: source
        // second index: destination
        // value: list of path distances from source to destination
        private Dictionary<AST.Address, Dictionary<AST.Address, HashSet<int>>> _matrix;
        private int _numVertices;

        public SparseMatrix(int numVertices)
        {
            _numVertices = numVertices;
            _matrix = new Dictionary<AST.Address, Dictionary<AST.Address, HashSet<int>>>(numVertices);
        }

        public void Connect(AST.Address source, AST.Address destination, int distance)
        {
            if (!_matrix.ContainsKey(source))
            {
                _matrix.Add(source, new Dictionary<AST.Address, HashSet<int>>());
            }

            if (!_matrix[source].ContainsKey(destination))
            {
                _matrix[source].Add(destination, new HashSet<int>());
            }

            _matrix[source][destination].Add(distance);
        }

        public SparseMatrix Transpose()
        {
            var transpose = new SparseMatrix(_numVertices);

            foreach (var kvp in _matrix)
            {
                var dest = kvp.Key;
                var sources = kvp.Value;

                foreach (var skvp in sources)
                {
                    var source = skvp.Key;
                    var distances = skvp.Value;

                    foreach (var distance in distances)
                    {
                        transpose.Connect(source, dest, distance);
                    }
                }
            }

            return transpose;
        }

        public HashSet<int> Distances(AST.Address from, AST.Address to)
        {
            return _matrix[from][to];
        }

        public Dictionary<AST.Address,HashSet<int>> AllRefDistances(AST.Address from)
        {
            return _matrix[from];
        }
    }
}
