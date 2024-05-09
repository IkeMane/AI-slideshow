#test Ai


from functions import generate_response


def test_ai():
    prompt = "What is the capital of Nigeria?"
    response = generate_response(prompt)
    print(response)


if __name__ == "__main__":
    test_ai()
