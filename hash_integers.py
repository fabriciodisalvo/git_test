"""
Task
----
Given an integer, n, and n space-separated integers as input,
create a tuple, t, of those n integers.
Then compute and print the result of hash(t).

Input Format
------------
The first line contains an integer, n, denoting the number
of elements in the tuple.
The second line contains n space-separated integers describing
the elements in tuple t.

Output Format
-------------
Print the result of hash(t).

Sample Input 0
--------------
2
1 2

Sample Output 0
---------------
3713081631934410656

#    n = int(input())
#    print(hash(tuple([int(x) for x in input().split()])))
"""


if __name__ == '__main__':
    n = int(input())
    integer_list = map(int, input().split())
    new_integer_list = list(integer_list)
    a, b = new_integer_list[0], new_integer_list[1]
    t = a, b
    hash_t = hash(t)
    print(hash_t)
