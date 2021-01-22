# Demo: global and local variables in VBA

Say I have two buttons on a UserForm. One declares and sets a variable, and the other reports the value of that variable to a `MsgBox`. Consider the following code.

```
Private Sub SetButton_Click()
    Dim message As String
    message = "Use the force, Luke"
End Sub

Private Sub SayButton_Click()
    MsgBox (message)
End Sub
```

This doesn't work. The problem here is that the variable `message` is defined and used in the `SetButton_Click()` sub only. When we get to `SayButton_Click()`, that sub doesn't know a variable called `message`.

The solution is to define `message` as a global variable, meaning it is available everywhere. In VBA, global variables are called 'Public', so the above code could be adapted as follows. Since `message` is defined as `Public` outside of either sub, it is available to both as a variable.

```
Public message As String

Private Sub SetButton_Click()
    message = "Use the force, Luke"
End Sub

Private Sub SayButton_Click()
    MsgBox (message)
End Sub
```

Here you will find two files.

- `Say_and_Say_Doesnt_work.xlsm` contains the first block of code. Open this and press to launch the UserForm. Press 'Set' to set the `message`, then press `Say` to report `message` to a `MsgBox`. You should find it is blank, because it does not know the variable `message`.
- `Say_and_Say_Public_variable.xlsm` contains the second block of code. Open this and press to launch the UserForm. Press 'Set' to set the `message`, then press `Say` to report `message` to a `MsgBox`. You should find it displays the string set in the first sub.

(Aside: if you set `Option Explicit`, VBA will not let you run the first example, reporting an error 'Variable not defined' on the 'Say' button, which ought to make it clear this is a problem.)
