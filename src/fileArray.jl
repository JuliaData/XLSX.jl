# Based on DiskArrays code suggested by Fabian Gans in
# https://discourse.julialang.org/t/struggling-to-use-mmap-with-ziparchives/129839/19
#
# Not `Mmap.mmap`: on Windows a mapped file can't be deleted or overwritten until GC
# finalizes the mapping, so a user couldn't rewrite a file they had just read (still the
# case on Julia 1.13). Instead the file is read in blocks, each with its own open and
# close, so it is never held open between reads. ZipArchives reads an entry in chunks of
# a few KiB, and on Windows an open per chunk cost more than the read itself (#482).
#
# A block starts at FILEARRAY_FIRST_BLOCK bytes and doubles with each refill up to
# FILEARRAY_BLOCK, so a short read, such as the central directory or a small part like
# sharedStrings.xml, costs a short read, while a long sequential one still reads 1 MiB
# at a time.
#
# ZipArchives makes a `view` per entry, in the task that reads it, and each view caches
# its own block, so tasks reading different entries don't evict each other's blocks.
# The lock covers the shared root array, whose headers several tasks may read at once.
const FILEARRAY_FIRST_BLOCK = 1 << 16
const FILEARRAY_BLOCK = 1 << 20

function FileArray(filename::String, offset::Int64, len::Int64)
    FileArray(filename, offset, len, ReentrantLock(), Int64(-1), UInt8[])
end
function FileArray(filename::String, offset::Int64=Int64(0))
    len = filesize(filename)
    len ≥ 0 || error("filesize of $(repr(filename)) is negative")
    offset ≥ 0 || error("offset $(offset) is negative")
    offset ≤ len || error("offset $(offset) is larger than the filesize $(len)")
    FileArray(filename, offset, len - offset)
end
Base.size(s::FileArray) = (s.len,)
function Base.getindex(s::FileArray, i::Int)::UInt8
    copyto!(zeros(UInt8, 1), Int64(1), s, Int64(i), Int64(1))[1]
end
# Any integer range, so that a view is a FileArray on 32-bit too, where `a:b` is Int32.
function Base.view(s::FileArray, inds::UnitRange{<:Integer})::FileArray
    checkbounds(s, inds)
    FileArray(s.filename, s.offset + Int64(first(inds)) - Int64(1), Int64(length(inds)))
end

# Cache the block starting at position `pos` (0-based), twice the size of the last one.
# Call with `s.lock` held.
function _fill_block!(s::FileArray, pos::Int64)
    n = min(Int64(clamp(2 * length(s.block), FILEARRAY_FIRST_BLOCK, FILEARRAY_BLOCK)), s.len - pos)
    resize!(s.block, n)
    s.block_start = Int64(-1)       # invalid until the read succeeds
    open(s.filename) do io
        seek(io, s.offset + pos)
        nb = readbytes!(io, s.block, n)
        nb == n || error("short read")
    end
    s.block_start = pos
    return nothing
end

dest_types = if VERSION ≥ v"1.11"
    (Vector{UInt8}, Memory{UInt8},)
else
    (Vector{UInt8},)
end
for dest_type in dest_types
    @eval begin
        function Base.copyto!(dest::$dest_type, dstart::Int64, src::FileArray, sstart::Int64, n::Int64)
            iszero(n) && return dest
            n ≥ 0 || throw(ArgumentError("tried to copy n=$(n) elements, but n should be non-negative"))
            checkbounds(dest, dstart)
            checkbounds(src, sstart)
            checkbounds(dest, dstart + n - Int64(1))
            checkbounds(src, sstart + n - Int64(1))
            pos = sstart - Int64(1)
            @lock src.lock while n > 0
                if !(src.block_start ≤ pos < src.block_start + length(src.block))
                    _fill_block!(src, pos)
                end
                o = pos - src.block_start
                k = min(n, length(src.block) - o)
                copyto!(dest, dstart, src.block, o + Int64(1), k)
                dstart += k; pos += k; n -= k
            end
            return dest
        end
    end
end
