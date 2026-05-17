import anthropic

# pip install anthropic
# Set ANTHROPIC_API_KEY environment variable before running

client = anthropic.Anthropic()

def ask(prompt: str) -> str:
    """Send a prompt to Claude and stream the response. Returns full response text."""
    print("Claude: ", end="", flush=True)

    with client.messages.stream(
        model="claude-opus-4-7",
        max_tokens=64000,
        thinking={"type": "adaptive"},
        cache_control={"type": "ephemeral"},
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)

        message = stream.get_final_message()

    print(f"\n\n[output tokens: {message.usage.output_tokens}]")
    return "".join(
        block.text for block in message.content if block.type == "text"
    )


if __name__ == "__main__":
    print("Claude API — max_tokens: 64,000  (type 'quit' to exit)\n")
    while True:
        user_input = input("You: ").strip()
        if user_input.lower() in ("quit", "exit", "q"):
            break
        if not user_input:
            continue
        ask(user_input)
        print()
