In some cases, the order in which a calculation is performed can affect the return value of the formula, so it's important to understand how the order is determined and how you can change the order to obtain the results you want.

Calculation order

Formulas calculate values in a specific order. A formula in Excel always begins with an equal sign (=). Excel interprets the characters that follow the equal sign as a formula. Following the equal sign are the elements to be calculated (the operands), such as constants or cell references. These are separated by calculation operators. Excel calculates the formula from left to right, according to a specific order for each operator in the formula.

Operator precedence in Excel formulas

If you combine several operators in a single formula, Excel performs the operations in the order shown in the following table. If a formula contains operators with the same precedence—for example, if a formula contains both a multiplication and division operator—Excel evaluates the operators from left to right.

Operator

Description

: (colon)

(single space)

, (comma)

Reference operators

–

Negation (as in –1)

%

Percent

^

Exponentiation

* and /

Multiplication and division

+ and –

Addition and subtraction

&

Connects two strings of text (concatenation)

=
< >
<=
>=
<>

Comparison

Using parentheses in Excel formulas

To change the order of evaluation, enclose in parentheses the part of the formula to be calculated first. For example, the following formula produces 11 because Excel performs multiplication before addition. The formula multiplies 2 by 3 and then adds 5 to the result.

=5+2*3

In contrast, if you use parentheses to change the syntax, Excel adds 5 and 2 together and then multiplies the result by 3 to produce 21.

=(5+2)*3

In the following example, the parentheses that enclose the first part of the formula force Excel to calculate B4+25 first and then divide the result by the sum of the values in cells D5, E5, and F5.
