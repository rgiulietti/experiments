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

/**
 * This class implements a method to properly encode a list of arguments
 * to form a valid command line on Windows.
 *
 * It assumes that the decoding of the command line by the invoked program
 * adheres to the conventions described in
 * <a href="https://docs.microsoft.com/en-us/cpp/c-language/parsing-c-command-line-arguments">Parsing C command-line arguments</a>.
 *
 * <h3>Examples</h3>
 *
 * <p>To avoid confusion with the special role of the quote {@code "} and
 * the backslash {@code \} in Java literals,
 * in the examples below strings are instead enclosed in brackets,
 * which are <em>not</em> part of their content.
 * The quotes and backslashes retain their literal values.</p>
 *
 * <p>Always quoted, even when not strictly necessary.</p>
 * <pre>
 * command={[program], [abc], [ ab], [a b], [ab ]}
 * encoded=["program" "abc" " ab" "a b" "ab "]
 * </pre>
 *
 * <p>Empty argument; balanced and unbalanced quotes in arguments.</p>
 * <pre>
 * command={[program path], [], ["a b"], ["a b], [a" b], [a b"]}
 * encoded=["program path" "" "\"a b\"" "\"a b" "a\" b" "a b\""]
 * </pre>
 *
 * <p>Backslashes followed by non-quote.</p>
 * <pre>
 * command={[program path], [a\b], [a\\b]}
 * encoded=["program path" "a\b" "a\\b"]
 * </pre>
 *
 * <p>Backslashes followed by quote.</p>
 * <pre>
 * command={[program path], [a\"b], [a\\"b]}
 * encoded=["program path" "a\\\"b" "a\\\\\"b"]
 * </pre>
 *
 * <p>Trailing backslashes.</p>
 * <pre>
 * command={[program path], [a\], [a\\]}
 * encoded=["program path" "a\\" "a\\\\"]
 * </pre>
 *
 * <p>Lone whitespaces; other characters.</p>
 * <pre>
 * command={[program path], [  ], [<&^|%]}
 * encoded=["program path" "  " "<&^|%"]
 * </pre>
 */
public class WinCmdlineEncoder {

    private static final char QUOTATION_MARK = '"';
    private static final char REVERSE_SOLIDUS = '\\';
    private static final char SPACE = ' ';
    private static final char NUL = '\0';

    private final String[] command;
    private final StringBuilder sb = new StringBuilder();
    private String arg;
    private int i; // index of char yet to read
    private int c; // current char, or -1 if beyond last char of arg

    private WinCmdlineEncoder(String[] command) {
        this.command = command;
    }

    /**
     * <p>Encodes the elements in {@code command} to form a command line.</p>
     *
     * <p>The command line can be passed to {@code CreateProcess()}
     * provided the invoked program parses it as documented in the reference
     * given above.
     * The runtime of traditional C/C++ console programs recovers the original
     * elements in the {@code argv[]} array of the {@code main()} function.</p>
     *
     * <p>However, other runtimes might not parse the command line in this way.
     * Moreover, even a traditional console program could parse the command line
     * (as returned by the Windows {@code GetCommandLine()} function) in any
     * way it deems useful.</p>
     *
     * @param command a non-empty array of non-null strings, each one
     *                not containing the NULL character.
     * @return a command line to pass to {@code CreateProcess()}.
     * @throws NullPointerException If {@code command} or any of its elements
     * is {@code null}.
     * @throws IndexOutOfBoundsException If {@code command} is empty.
     * @throws IOException If any of the elements of {@code command} contains
     * a NUL character.
     */
    public static String encode(String[] command) throws IOException {
        return new WinCmdlineEncoder(command).encode();
    }

    private String encode() throws IOException {
        encodeFirstArg();
        for (int k = 1; k < command.length; ++k) {
            appendSpace();
            encodeArg(k);
        }
        return sb.toString();
    }

    /*
     * In the first arg (index 0), backslashes are normal characters and
     * quotes are not allowed.
     */
    private void encodeFirstArg() throws IOException {
        arg = command[0];
        i = 0;
        appendQuote(); // opening quote
        read();
        while (!isEos()) {
            if (isQuote()) {
                throw new IOException("invalid '\"' character in program");
            }
            append();
            read();
        }
        appendQuote(); // closing quote
    }

    /*
     * In other args, backslashes and quotes are encoded.
     */
    private void encodeArg(int k) throws IOException {
        arg = command[k];
        i = 0;
        appendQuote(); // opening quote
        read();
        while (!isEos()) {
            if (isBackslash()) {
                encodeBackslashes();
            } else if (isQuote()) {
                appendEncodedQuote();
            } else {
                append();
            }
            read();
        }
        appendQuote(); // closing quote
    }

    /*
     * n backslashes followed by a quote are encoded as 2 n + 1 literal \
     * and a literal ".
     *
     * n backslashes followed by a character other than a quote
     * are encoded as n literal \ and the character.
     *
     * n trailing backslashes are encoded as 2 n literal \
     */
    private void encodeBackslashes() throws IOException {
        var n = 0;
        while (isBackslash()) {
            ++n;
            read();
        }
        if (isQuote()) {
            appendEncodedBackslashes(n);
            appendEncodedQuote();
        } else if (isEos()) {
            appendEncodedBackslashes(n);
        } else {
            appendBackslashes(n);
            append();
        }
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

    private void appendEncodedBackslashes(int count) {
        for (; count > 0; --count) {
            appendBackslash();
            appendBackslash();
        }
    }

    private void appendBackslashes(int count) {
        for (; count > 0; --count) {
            appendBackslash();
        }
    }

    private void append() {
        sb.append((char) c);
    }

    private void appendSpace() {
        sb.append(SPACE);
    }

    private void appendBackslash() {
        sb.append(REVERSE_SOLIDUS);
    }

    private void appendQuote() {
        sb.append(QUOTATION_MARK);
    }

    private void appendEncodedQuote() {
        appendBackslash();
        appendQuote();
    }

    /*
     * Read next char (or -1 if at end) into c
     */
    private void read() throws IOException {
        c = i < arg.length() ? arg.charAt(i++) : -1;
        if (c == NUL) {
            throw new IOException("invalid NUL character in command");
        }
    }

}
