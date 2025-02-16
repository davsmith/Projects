<!-- Headings -->
# Heading 1

## Heading 2

### Heading 3

#### Heading 4

##### Heading 5

###### Heading 6

<!-- New line -->
To separate lines without an additional lines pace, end the line with two spaces and \<enter> or \<br>.
Like this.  
Also, the **the backslash (\\)** can be used as an escape character  
And this should be indented
<!-- Italics -->
*This text* is italic  
<!-- An underscore may also be used to demarcate italic text, but the linter will complain -->

<!-- Strong -->
**This text** is strong (bold)  
<!-- Two underscores may also be used to demarcate bold text, but the linter will complain -->

<!-- Strike through -->
~~This text~~ is strike through

<!-- Horizontal Rule -->

---

<!-- Blockquote -->
> Single line quote
>> Nested quote
>> multiple line
>> quote
>>> Even more nested quote
>> Back to the first nesting
> Top level quote

<!-- Links -->
[Microsoft](http://www.microsoft.com)  
[Microsoft (with hover text)](http://www.microsoft.com "This is the hover text")

<!-- UL -->
* Item 1
* Item 2
* Item 3
  * Nested Item 1
  * Nested Item 2

<!-- OL -->
1. Item 1
1. Item 2
1. Item 3

<!-- Inline Code Block -->
`<p>This is a paragraph</p>`

<!-- Images -->
![Markdown Logo](https://uhf.microsoft.com/images/microsoft/RE1Mu3b.png "Images also have hover text")

# Code Blocks
<!-- Code Blocks -->
```bash
  npm install

  npm start
```

```javascript
  function add(num1, num2) {
    return num1 + num2;
  }
```

```python
  def add(num1, num2):
    return num1 + num2
```

``` js
const count = records.length;
```

``` csharp
Console.WriteLine("Hello, World!");
```

<!-- Tables -->
| Name     | Email          |Right Justified|Center|
| -------- | -------------- |--------------:|:----:|
| John Doe | <john@gmail.com> |Notes 1|Text 1|
| Jane Doe | <jane@gmail.com> |Notes 2|Text 2|

<!-- Task List -->
* [x] Task 1
* [x] Task 2
* [ ] Task 3

# HTML Elements

The MarkDown linter doesn't appreciate HTML elements, but some are handy  

<font color="red">Red text</font> for example  
<font color="blue">Blue text</font> is also sometimes handy

<p>This text needs <del>strike through</del> and this text should be <ins>underlined</ins>
<p><tt>This text is teletype text.</tt></p>
<center>This text is center-aligned.</center>
<p>This text contains <sup>superscript</sup> text.</p>
<p>This text contains <sub>subscript</sub> text.</p>
<p>The project status is <span style="color:green;font-weight:bold">GREEN</span> even though the bug count / developer may be in <span style="color:red;font-weight:bold">red.</span> - Capability of span
<p><small>Disclaimer: Wiki also supports showing small text</small></p>
<p><big>Bigger text</big></p>
This is a normal line of text.  Nothing special.  
<br><br><br>
<details>
  <summary>Expandable Text</summary>  

## Heading

  1. A numbered
  2. list
     * With some
     * Sub bullets  

</details>

[Jump to a Section Title](#blibbit)

# Symbols

$
\alpha, \beta, \gamma, \delta, \epsilon, \zeta, \eta, \theta, \kappa, \lambda, \mu, \nu, \omicron, \pi, \rho, \sigma, \tau, \upsilon, \phi, ...
$  

$\Gamma,  \Delta,  \Theta, \Lambda, \Xi, \Pi, \Sigma, \Upsilon, \Phi, \Psi, \Omega$
<br><br><br>

# Mermaid

::: mermaid
gantt
    title A Gantt chart
    dateFormat YYYY-MM-DD
    excludes 2022-03-16,2022-03-18,2022-03-19
    section Section
    A task          :a1, 2022-03-012, 7d
    Another task    :after a1 , 5d
:::

<a id="blibbit"></a>  
This should have an anchor  
There is more text here  
Don't forget me!  
