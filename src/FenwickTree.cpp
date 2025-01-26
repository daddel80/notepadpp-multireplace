#include "FenwickTree.h"
#include <Windows.h>

// Constructor: Initialize with 0 size.
FenwickTree::FenwickTree() : _size(0) {}

// Initialize the Fenwick Tree with a given size (1-based indexing).
void FenwickTree::init(size_t n) {
    _size = n;
    fenwicksum.assign(n + 1, 0); // 1 extra for 1-based indexing
}

// Update: Add 'delta' to all elements >= index (1-based indexing).
void FenwickTree::update(size_t index, LRESULT delta) {
    while (index <= _size) {
        fenwicksum[index] += delta;
        index += static_cast<size_t>(index & -static_cast<std::make_signed_t<size_t>>(index)); // move to the next node
    }
}

// Query: Get the prefix sum from 1 to index (1-based indexing).
LRESULT FenwickTree::prefixSum(size_t index) const {
    LRESULT result = 0;
    while (index > 0) {
        result += fenwicksum[index];
        index -= static_cast<size_t>(index & -static_cast<std::make_signed_t<size_t>>(index)); // move to the parent node
    }
    return result;
}

// Get the size of the Fenwick Tree.
size_t FenwickTree::size() const {
    return _size;
}
