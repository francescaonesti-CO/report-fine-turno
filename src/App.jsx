import { useState } from "react";
import { supabase } from "./lib/supabaseClient";

export default function App() {
  const [form, setForm] = useState({
    operator_name: "",
    notes: "",
  });

  const handleChange = (e) => {
    setForm({ ...form, [e.target.name]: e.target.value });
  };

  const salvaReport = async () => {
    const { error } = await supabase.from("reports").insert([
      {
        command_id: "ae6f07c1-404f-41a1-9be7-9ff0bc83c325",
        service_date: new Date().toISOString().split("T")[0],
        status: "inviato",
        notes: form.notes,
      },
    ]);

    if (error) {
      alert("Errore salvataggio");
      console.error(error);
    } else {
      alert("Report salvato!");
      setForm({ operator_name: "", notes: "" });
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>Report Fine Turno</h2>

      <input
        name="operator_name"
        placeholder="Operatore"
        value={form.operator_name}
        onChange={handleChange}
      />

      <br /><br />

      <textarea
        name="notes"
        placeholder="Note"
        value={form.notes}
        onChange={handleChange}
      />

      <br /><br />

      <button onClick={salvaReport}>Salva Report</button>
    </div>
  );
}
