/*
 * Copyright (c) 2022 Raffaello Giulietti
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

import java.io.IOException;
import java.util.ArrayList;

/**
 * <p>This class implements a more sophisticated algorithm to parse the
 * command line on Windows than the one currently used.</p>
 *
 * <p>{@code Runtime.exec(String)} and related methods that accept a single
 * string as the command line are {@code @Deprecated(since="18")} but are
 * not yet marked {@code @Deprecated(forRemoval=true)}.
 * The parsing of the command string is based on "whitespace" as separator.
 * While pretty naive on all Unix-like and Windows platforms, this parsing
 * is at least clear and simple to explain and works in many cases.
 * Its simplicity, though, causes unintended misbehavior in other cases.
 * Moreover, since there's no parameter akin to the application name of
 * {@code CreateProcess()} (see below), the program must be identified
 * in the initial section of the command line.
 * Finally, "whitespace" as used in these methods is more comprehensive than
 * the one used in the Windows parsing, potentially adding more trouble.</p>
 *
 * <p>On Windows, a process is created and started in one go by invoking
 * the {@code CreateProcess()} function.
 * It accepts an application name and a <em>single</em> command line string,
 * <em>freely</em> encoding the arguments.</p>
 *
 * <p>There are two parsing steps performed when a new process is created.
 * The first is performed by {@code CreateProcess()} itself in the invoking
 * program, while the second one is done in the new program.</p>
 *
 * <p>{@code CreateProcess()} needs to identify which program to launch.
 * It does so in a rather peculiar way described in
 * <a href="https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessa">CreateProcess function</a>.</p>
 *
 * <p>Among other rules, if the application name is specified, it is parsed;
 * otherwise, the initial section of the command line is parsed.
 * The convention is that if the application name is specified, then it should
 * be repeated at the beginning of the command line as well.
 * The parsing combines lexical rules with a search path and some checks
 * on the presence of files on the filesystem.
 * The important point, here, is that this parsing is not merely based
 * on lexical rules but is driven by the content of the filesystem as well.</p>
 *
 * <p>The documentation recommends quoting the program on the command line
 * to resolve any ambiguity arising from the parsing strategy on unquoted
 * programs, so as to avoid launching an unintended one.</p>
 *
 * <p>Once the program is identified in either way, the rest of the command line
 * is <em>not</em> parsed by {@code CreateProcess()} but by the new program.
 * The new program can obtain the command line by invoking the function
 * {@code GetCommandLine()}.</p>
 *
 * <p>There's no hard set of rules on how the command line string is parsed
 * into tokens and how these are interpreted.
 * A program chooses any strategy deemed useful to make sense of the string.</p>
 *
 * <p>However, if the invoked program is a traditional C/C++ application,
 * the language runtime populates the {@code argv[]} array in {@code main()}
 * by parsing the command line according to rules described in
 * <a href="https://docs.microsoft.com/en-us/cpp/c-language/parsing-c-command-line-arguments">Parsing C command-line arguments</a>.</p>
 *
 * <p>The initial section of the command line is possibly parsed twice.
 * Once by {@code CreateProcess()} if the application name is left unspecified,
 * and once by the runtime.
 * The results might thus be different.</p>
 *
 * <p>The algorithm implemented by this class attempts to mimic the parsing
 * performed by the C/C++ runtime to populate {@code argv[]}.
 * However, mimicking the parsing of {@code CreateProcess()} to determine
 * the program from the command line by querying the filesystem is deemed too
 * cumbersome and is not done here.
 * Rather, the method assumes that the program is fully specified on
 * the command line, preferably in quoted form.</p>
 *
 *
 * <h3>Examples</h3>
 *
 * <p>To avoid confusion with the special role of the quote {@code "} and
 * {@code \} in both Java literals and in command lines, in the examples below
 * strings are instead enclosed in brackets, which are <em>not</em> part
 * of their content.</p>
 *
 * <p>Business as usual, but the program must always be quoted.</p>
 * <pre>
 * command=["C:\Windows\System32\cmd.exe" /c batch.bat]
 * program=[C:\Windows\System32\cmd.exe]
 * args[0]=[C:\Windows\System32\cmd.exe]
 * args[1]=[/c]
 * args[2]=[batch.bat]
 * </pre>
 *
 * <p>Quotes are used to include whitespace in arguments.
 * Trailing whitespaces outside quotes are skipped.</p>
 * <pre>
 * command=["C:\Program Files\program path.exe" "some text.txt"  ]
 * program=[C:\Program Files\program path.exe]
 * args[0]=[C:\Program Files\program path.exe]
 * args[1]=[some text.txt]
 * </pre>
 *
 * <p>The same outcome with a more cumbersome way to quote.</p>
 * <pre>
 * command=["C:\Program Files\program path.exe" som"e "tex""t.t"x"t  ]
 * program=[C:\Program Files\program path.exe]
 * args[0]=[C:\Program Files\program path.exe]
 * args[1]=[some text.txt]
 * </pre>
 *
 * <p>The following example illustrates the difference between the program
 * and the first argument.
 * It might seem weird, but this is the way Windows works!</p>
 * <pre>
 * command=["C:\Program Files\program path.exe"som"e "tex""t.t"x"t   ]
 * program=[C:\Program Files\program path.exe]
 * args[0]=[C:\Program Files\program path.exesome text.txt]
 * </pre>
 *
 * <p>Backslashes and quotes work as in
 * <a href="https://docs.microsoft.com/en-us/cpp/c-language/parsing-c-command-line-arguments">Parsing C command-line arguments</a>.
 * Note that the last argument is not properly closed by a quote,
 * so the trailing whitespaces become part of it.</p>
 * <pre>
 * command=["program" \" \\\" "" "C:\\" """Hi!"", said the guy, \"Hi!\"  ]
 * program=[program]
 * args[0]=[program]
 * args[1]=["]
 * args[2]=[\"]
 * args[3]=[]
 * args[4]=[C:\]
 * args[5]=["Hi!", said the guy, "Hi!"  ]
 * </pre>
 *
 * <p>Almost the same command with a backslash removed.
 * Quite a different outcome!</p>
 * <pre>
 * command=["program" \" \\" "" "C:\\" """Hi!"", said the guy, \"Hi!\" ]
 * program=[program]
 * args[0]=[program]
 * args[1]=["]
 * args[2]=[\ " C:\ "Hi!,]
 * args[3]=[said]
 * args[4]=[the]
 * args[5]=[guy,]
 * args[6]=["Hi!"]
 * </pre>
 */
public class WinCommandParser {

    public record WinCommand(String program, String[] args) {}

    private static final char QUOTATION_MARK = '"';
    private static final char REVERSE_SOLIDUS = '\\';
    private static final char SPACE = ' ';
    private static final char HT = '\t';
    private static final char NUL = '\0';

    private final String command;
    private final StringBuilder arg = new StringBuilder();
    private final ArrayList<String> args = new ArrayList<>();
    private String program;
    private int i; // index of char yet to read
    private int c; // current char, or -1 if beyond last char of command

    /**
     * Parses the supplied {@code command}.
     *
     * @param command A command line to be parsed.
     * @return The command line split in its constituents by the parsing.
     * @throws NullPointerException If {@code command} is {@code null}.
     * @throws IOException If {@code command} contains a NULL character.
     */
    public static WinCommand parse(String command) throws IOException {
        var parser = new WinCommandParser(command);
        parser.parse();
        return new WinCommand(parser.program, parser.args());
    }

    private WinCommandParser(String command) {
        this.command = command;
    }

    private String[] args() {
        return args.toArray(new String[args.size()]);
    }

    private void parse() throws IOException {
        read();
        parseFirstArg();
        while (!isEos()) {
            parseArg();
        }
    }

    /*
     * The first arg (index 0) is parsed differently than the others.
     * Backslashes and adjacent quotes have no special meaning.
     *
     * WARNING: This aspect of parsing is not well documented by Microsoft and
     * has been determined mostly experimentally on Windows.
     * Care is advised.
     *
     * Contrary to what is described in
     * <a href="https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessa">CreateProcess function</a>,
     * no attempt is made to parse for a match on the filesystem
     * by growing the program string because the first part must be quoted.
     *
     * The first arg is a sequence of plain quoted and plain unquoted parts
     * starting with a plain quoted part (see parsePlainQuotedPart() and
     * parsePlainUnquotedPart()).
     * The first unquoted whitespace or the end of the command,
     * whichever comes first, ends the arg.

     * The first (and possibly only) part designates the program,
     * but the arg itself might include other parts as well.
     * See the class documentation for an unusual example.
     */
    private void parseFirstArg() throws IOException {
        /* the first part denotes the program... */
        if (!isQuote()) {
            throw new IOException("the program path must be quoted");
        }
        parsePlainQuotedPart();
        program = arg.toString();

        /* ... but the first arg can consist of more parts */
        while (!isWhitespace() && !isEos()) {
            if (isQuote()) {
                parsePlainQuotedPart();
            } else {
                parsePlainUnquotedPart();
            }
        }
        addArg();
        skipWhitespace();
    }

    /*
     * An arg other than the first is a sequence of quoted and unquoted parts.
     * (See parseQuotedPart() and parseUnquotedPart().)
     * The first unquoted whitespace or the end of the command,
     * whichever comes first, ends the arg.
     */
    private void parseArg() throws IOException {
        while (!isWhitespace() && !isEos()) {
            if (isQuote()) {
                parseQuotedPart();
            } else {
                parseUnquotedPart();
            }
        }
        addArg();
        skipWhitespace();
    }

    /*
     * A plain quoted part spans from the current quote to the next quote
     * or the end of the command, whichever comes first.
     * The quoted part, without the enclosing quotes, is appended to arg.
     */
    private void parsePlainQuotedPart() throws IOException {
        read(); // consume opening quote
        while (!isQuote() && !isEos()) {
            appendRead();
        }
        read(); // consume closing quote: harmless when isEos()
    }

    /*
     * A plain unquoted part spans from the current char to the next whitespace,
     * quote or the end of the command, whichever comes first.
     * The unquoted part, without the whitespace or quote, is appended to arg.
     */
    private void parsePlainUnquotedPart() throws IOException {
        while (!isWhitespace() && !isQuote() && !isEos()) {
            appendRead();
        }
    }

    /*
     * A quoted part spans from the current quote to the next quote
     * or the end of the command, whichever comes first.
     * But neither quote of each pair of adjacent quotes (after the opening
     * quote) closes the quoted part. Rather, for each such pair, a literal "
     * is appended to it.
     * Sequences of one or more backslashes are processed separately.
     * The quoted part, without the enclosing quotes, is appended to arg.
     */
    private void parseQuotedPart() throws IOException {
        read(); // consume opening quote
        for (;;) { // this form best conveys the existence of more exit points
            if (isQuote()) {
                read(); // consume potentially closing quote
                if (isQuote()) {
                    appendRead(); // another quote, append and consume
                } else {
                    break; // a closing quote, indeed: already consumed above
                }
            } else if (isBackslash()) {
                parseBackslashes();
            } else if (isEos()) {
                break;
            } else {
                appendRead();
            }
        }
    }

    /*
     * An unquoted part spans from the current char to the next whitespace,
     * quote or the end of the command, whichever comes first.
     * Sequences of one or more backslashes are processed separately.
     * The unquoted part, without the whitespace or quote, is appended to arg.
     */
    private void parseUnquotedPart() throws IOException {
        while (!isWhitespace() && !isQuote() && !isEos()) {
            if (isBackslash()) {
                parseBackslashes();
            } else {
                appendRead();
            }
        }
    }

    /*
     * A sequence of n = 2 k (even) backslashes and a quote means appending
     * k literal \ to the arg.
     * The quote itself opens or closes a quoted part.
     *
     * A sequence of n = 2 k + 1 (odd) backslashes and a quote means appending
     * k literal \ to the arg and a literal ".
     * The quote itself is thus consumed in this method.
     *
     * A sequence of n backslashes not followed by a quote means appending
     * n literal \ to the arg.
     */
    private void parseBackslashes() throws IOException {
        var n = 0;
        while (isBackslash()) {
            ++n;
            read();
        }
        if (isQuote()) {
            appendBackslashes(n / 2);
            if (n % 2 != 0) {
                appendRead(); // odd n, append and consume the quote
            }
        } else {
            appendBackslashes(n);
        }
    }

    private void skipWhitespace() throws IOException {
        while (isWhitespace()) {
            read();
        }
    }

    private void addArg() {
        args.add(arg.toString());
        arg.setLength(0);
    }

    /*
     * Read next char (or -1 if at end) into c
     */
    private void read() throws IOException {
        c = i < command.length() ? command.charAt(i++) : -1;
        if (c == NUL) {
            throw new IOException("invalid NULL character in command");
        }
    }

    private boolean isWhitespace() {
        return c == SPACE || c == HT;
    }

    private boolean isQuote() {
        return c == QUOTATION_MARK;
    }

    private boolean isBackslash() {
        return c == REVERSE_SOLIDUS;
    }

    private boolean isEos() {
        return c < 0;
    }

    private void appendRead() throws IOException {
        arg.append((char) c);
        read();
    }

    private void appendBackslashes(int count) {
        for (; count > 0; --count) {
            arg.append(REVERSE_SOLIDUS);
        }
    }

}
