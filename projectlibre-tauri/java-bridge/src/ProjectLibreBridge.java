import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;

public final class ProjectLibreBridge {
    private static final class BridgeState {
        private final List<String> sampleFiles = new ArrayList<String>();
        private String openedProject = "";
        private String lastCommand = "";

        private void setSampleFiles(String[] args) {
            sampleFiles.clear();
            for (String arg : args) {
                if (arg != null && arg.length() > 0) {
                    sampleFiles.add(arg);
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        BridgeState state = new BridgeState();
        state.setSampleFiles(args);

        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in, StandardCharsets.UTF_8));
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(System.out, StandardCharsets.UTF_8));

        writer.write("{\"ok\":true,\"command\":\"ready\",\"data\":{\"message\":\"ProjectLibre Java bridge ready\"}}\n");
        writer.flush();

        String line;
        while ((line = reader.readLine()) != null) {
            if (line.trim().length() == 0) {
                continue;
            }

            BridgeResponse response = handle(line, state);
            writer.write(response.toJson());
            writer.write('\n');
            writer.flush();

            if (response.quit) {
                break;
            }
        }
    }

    private static BridgeResponse handle(String line, BridgeState state) {
        state.lastCommand = line;
        String[] parts = line.split("\t");
        String command = parts.length > 0 ? parts[0] : "";

        if ("ping".equals(command)) {
            return BridgeResponse.ok("ping", jsonMessage("bridge ready"));
        }
        if ("snapshot".equals(command)) {
            return BridgeResponse.ok("snapshot", buildSnapshot(state));
        }
        if ("open".equals(command) || "import_mpp".equals(command)) {
            String path = parts.length > 1 ? decode(parts[1]) : "";
            state.openedProject = path;
            return BridgeResponse.ok(command, buildOpenState(state, path));
        }
        if ("export_mpp".equals(command)) {
            String target = parts.length > 1 ? decode(parts[1]) : "";
            if (state.openedProject.length() == 0) {
                return BridgeResponse.error(command, "No project is open");
            }
            return BridgeResponse.ok(command, buildExportState(state, target));
        }
        if ("quit".equals(command)) {
            return BridgeResponse.quit(command, jsonMessage("bridge shutting down"));
        }

        return BridgeResponse.error(command, "Unknown command: " + command);
    }

    private static String buildSnapshot(BridgeState state) {
        StringBuilder out = new StringBuilder();
        out.append("{");
        out.append("\"opened_project\":").append(jsonString(state.openedProject.length() == 0 ? null : state.openedProject)).append(",");
        out.append("\"last_command\":").append(jsonString(state.lastCommand)).append(",");
        out.append("\"sample_files\":").append(jsonArray(state.sampleFiles)).append(",");
        out.append("\"bridge_mode\":\"java-wrapper\"");
        out.append("}");
        return out.toString();
    }

    private static String buildOpenState(BridgeState state, String path) {
        StringBuilder out = new StringBuilder();
        out.append("{");
        out.append("\"opened_project\":").append(jsonString(path)).append(",");
        out.append("\"opened_name\":").append(jsonString(basename(path))).append(",");
        out.append("\"sample_files\":").append(jsonArray(state.sampleFiles));
        out.append("}");
        return out.toString();
    }

    private static String buildExportState(BridgeState state, String target) {
        StringBuilder out = new StringBuilder();
        out.append("{");
        out.append("\"source_project\":").append(jsonString(state.openedProject)).append(",");
        out.append("\"target_path\":").append(jsonString(target)).append(",");
        out.append("\"opened_name\":").append(jsonString(basename(state.openedProject)));
        out.append("}");
        return out.toString();
    }

    private static String jsonMessage(String message) {
        return "{\"message\":" + jsonString(message) + "}";
    }

    private static String jsonArray(List<String> values) {
        StringBuilder out = new StringBuilder();
        out.append("[");
        for (int index = 0; index < values.size(); index++) {
            if (index > 0) {
                out.append(",");
            }
            out.append(jsonString(values.get(index)));
        }
        out.append("]");
        return out.toString();
    }

    private static String jsonString(String value) {
        if (value == null) {
            return "null";
        }
        StringBuilder out = new StringBuilder();
        out.append("\"");
        for (int index = 0; index < value.length(); index++) {
            char ch = value.charAt(index);
            switch (ch) {
                case '\\':
                    out.append("\\\\");
                    break;
                case '\"':
                    out.append("\\\"");
                    break;
                case '\b':
                    out.append("\\b");
                    break;
                case '\f':
                    out.append("\\f");
                    break;
                case '\n':
                    out.append("\\n");
                    break;
                case '\r':
                    out.append("\\r");
                    break;
                case '\t':
                    out.append("\\t");
                    break;
                default:
                    if (ch < 0x20) {
                        out.append(String.format("\\u%04x", (int) ch));
                    } else {
                        out.append(ch);
                    }
                    break;
            }
        }
        out.append("\"");
        return out.toString();
    }

    private static String basename(String path) {
        if (path == null || path.length() == 0) {
            return "";
        }
        return new File(path).getName();
    }

    private static String decode(String value) {
        if (value == null || value.length() == 0) {
            return "";
        }
        return new String(Base64.getDecoder().decode(value), StandardCharsets.UTF_8);
    }

    private static final class BridgeResponse {
        private final boolean ok;
        private final boolean quit;
        private final String command;
        private final String payload;
        private final String error;

        private BridgeResponse(boolean ok, boolean quit, String command, String payload, String error) {
            this.ok = ok;
            this.quit = quit;
            this.command = command;
            this.payload = payload;
            this.error = error;
        }

        private static BridgeResponse ok(String command, String payload) {
            return new BridgeResponse(true, false, command, payload, null);
        }

        private static BridgeResponse error(String command, String error) {
            return new BridgeResponse(false, false, command, null, error);
        }

        private static BridgeResponse quit(String command, String payload) {
            return new BridgeResponse(true, true, command, payload, null);
        }

        private String toJson() {
            StringBuilder out = new StringBuilder();
            out.append("{");
            out.append("\"ok\":").append(ok ? "true" : "false").append(",");
            out.append("\"command\":").append(jsonString(command));
            if (payload != null) {
                out.append(",\"data\":").append(payload);
            }
            if (error != null) {
                out.append(",\"error\":").append(jsonString(error));
            }
            out.append("}");
            return out.toString();
        }
    }
}
