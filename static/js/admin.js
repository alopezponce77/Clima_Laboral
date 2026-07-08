document.querySelectorAll(".edit-user").forEach(button => {
  button.addEventListener("click", () => {
    document.getElementById("user_id").value = button.dataset.id;
    document.getElementById("nombre").value = button.dataset.nombre;
    document.getElementById("usuario").value = button.dataset.usuario;
    document.getElementById("activo").checked = button.dataset.activo === "1";
    document.getElementById("password").placeholder = "Dejar vacia para conservar";
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});

document.querySelectorAll(".edit-row").forEach(button => {
  button.addEventListener("click", () => {
    document.getElementById("row_id").value = button.dataset.id;
    document.getElementById("activo").checked = button.dataset.activo === "1";
    Object.keys(button.dataset).forEach(key => {
      if (key === "activo") return;
      const input = document.getElementById(key);
      if (input) input.value = button.dataset[key];
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});
