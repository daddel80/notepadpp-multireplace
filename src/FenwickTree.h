#ifndef FENWICKTREE_H
#define FENWICKTREE_H

#include <vector>
#include <cstddef> // for size_t
#include <Windows.h>

// FenwickTree (Binary Indexed Tree) for efficient prefix sum calculations.
class FenwickTree {
private:
    std::vector<LRESULT> fenwicksum;
    size_t _size;

public:
    FenwickTree();
    void init(size_t n);
    void update(size_t index, LRESULT delta);
    LRESULT prefixSum(size_t index) const;
    size_t size() const;
};

#endif // FENWICKTREE_H
