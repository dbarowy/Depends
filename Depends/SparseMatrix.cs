using System;
using System.Collections.Generic;

namespace Depends
{
    [Serializable]
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

        public SparseMatrix(SparseMatrix other)
        {
            _numVertices = other._numVertices;
            _matrix = new Dictionary<AST.Address, Dictionary<AST.Address, HashSet<int>>>(_numVertices);

            foreach (var kvp in other._matrix)
            {
                var source = kvp.Key;
                var dests = kvp.Value;
                foreach (var dkvp in dests)
                {
                    var dest = dkvp.Key;
                    var distances = dkvp.Value;

                    foreach (var distance in distances)
                    {
                        this.Connect(source, dest, distance);
                    }
                }
            }
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
            if (!_matrix.ContainsKey(from))
            {
                return new HashSet<int>();
            } else if (!_matrix[from].ContainsKey(to))
            {
                return new HashSet<int>();
            } else
            {
                return _matrix[from][to];
            }
        }

        public Dictionary<AST.Address,HashSet<int>> AllRefDistances(AST.Address from)
        {
            if (!_matrix.ContainsKey(from))
            {
                return new Dictionary<AST.Address, HashSet<int>>();
            } else
            {
                return _matrix[from];
            }
        }
    }
}
